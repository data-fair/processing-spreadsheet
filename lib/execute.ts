import type { RunFunction } from '@data-fair/lib-common-types/processings.js'
import type { ProcessingConfig, CreateDatasets, UpdateDatasets, Parameters } from '#types/processingConfig/index.ts'
import { formatBytes } from '@data-fair/lib-utils/format/bytes.js'

import util from 'util'
import fs from 'fs-extra'
import path from 'path'
import Excel from 'exceljs'
import FormData from 'form-data'

import type { SheetsTab, SpreadsheetProcessingContext } from './context.ts'
import { fetchHTTP } from './fetch.ts'
import { createTmpFile } from './tmp-file.ts'
import { runCommand } from './spawn-process.ts'

/**
 * Allows for a requested program shutdown to be scheduled.
 */
let shouldBeStopped = false

export const stop: () => Promise<void> = async () => {
  shouldBeStopped = true
}

type SheetsList = Record<number, { name: string, featureCount: number }>

/**
 * Input function, allows data processing to begin
 * @param context Context of the request
 */
export const run: RunFunction<ProcessingConfig> = async (context) => {
  shouldBeStopped = false

  // Retrieving the contextual elements necessary for processing
  const { processingConfig, patchConfig } = context
  const tmpFile = await download(context)

  if (shouldBeStopped) return
  if (!tmpFile) return
  const sheetsList = await extraction(context, tmpFile)

  if (shouldBeStopped) return
  if (!sheetsList) return

  if (processingConfig.datasetMode === 'create') {
    const result = await createDatasets(context, sheetsList, tmpFile)

    // The lib-common-types signature only allows `dataset` (singular), but the worker's
    // patchConfig is a generic Object.assign on the config — `datasets` is supported at runtime.
    if (result?.updateConfig?.length) await patchConfig({ datasetMode: 'update', datasets: result.updateConfig, sheets: result.sheetsTab, url: processingConfig.url.trim() } as any)
  } else if (processingConfig.datasetMode === 'update') {
    await updateDatasets(context, sheetsList, tmpFile)
    await patchConfig({ url: processingConfig.url.trim() } as any)
  } else {
    const createConfig: SheetsTab[] = Object.keys(sheetsList).map(idSheet => ({
      add: false,
      nb: Number(idSheet),
      name: sheetsList[Number(idSheet)].name,
      lines: sheetsList[Number(idSheet)].featureCount,
      titleEditable: '',
      titleReadOnly: ''
    }))

    await patchConfig({ datasetMode: 'create', haveList: true, sheets: createConfig, dataset: {}, url: processingConfig.url.trim() } as any)
  }
}

/**
 * Allows you to download the file and place it in a temporary folder for later processing.
 * We only process .zip and .xlsx formats; any other format will result in an error.
 * @param processingConfig  Processing configuration, obtained from the form data (processing-config-schema.json)
 * @param tmpDir            Directory where to download the file
 * @param axios             Server for API requests
 * @param log               Log system that is displayed on the user interface
 * @returns Full path of the file to be processed
 */
const download = async ({ processingConfig, tmpDir, axios, log } : SpreadsheetProcessingContext) => {
  await fs.ensureDir(tmpDir)

  await log.step('Téléchargement du fichier')
  let tmpFile = path.join(tmpDir, 'file')
  await fs.ensureFile(tmpFile)
  if (shouldBeStopped) return

  const url = new URL(processingConfig.url.trim())
  let filename = decodeURIComponent(path.basename(url.pathname))
  if (shouldBeStopped) return

  filename = await fetchHTTP(processingConfig, tmpFile, axios) || filename
  if (shouldBeStopped) return

  // Try to prevent weird bug with NFS by forcing syncing file before reading it
  const fd = await fs.open(tmpFile, 'r')
  await fs.fsync(fd)
  await fs.close(fd)
  await log.info(`Le fichier a été téléchargé (${filename})`)
  if (shouldBeStopped) return

  let xlsxFilename

  // Check the file format
  if (filename.toLowerCase().endsWith('.zip')) {
    await log.info(`Dézippage du fichier ${filename}`)

    // Unzip
    await runCommand('unzip', ['-j', tmpFile, '-d', `${tmpFile}-dezip`])

    // We are looking for the .xlsx files contained in the .zip file.
    const filesXlsx: string[] = []
    const files = await fs.readdir(`${tmpFile}-dezip`)
    for (const file of files) {
      if (file.toLowerCase().endsWith('.xlsx')) {
        filesXlsx.push(`${tmpFile}-dezip/${file}`)
      }
    }

    const nbFiles = filesXlsx.length
    if (shouldBeStopped) return

    if (nbFiles <= 0) {
      throw new Error('Il n\'y a pas de fichiers .xlsx à traiter dans ce zip.')
    } else {
      // We keep the first .xlsx file we find, we ignore the others
      xlsxFilename = path.basename(filesXlsx[0])
      tmpFile = filesXlsx[0]
    }
  } else if (filename.toLowerCase().endsWith('.xlsx')) {
    xlsxFilename = filename
  } else {
    throw new Error('Le format n\'est pas pris en charge')
  }

  await log.info(`Traitement du fichier ${xlsxFilename}`)

  return tmpFile
}

/**
 * Allows you to retrieve the sheets of a file and organize their structure
 * @param log       Log system that is displayed on the user interface
 * @param tmpFile   Full path of the file to be processed
 * @returns Dictionary of available sheet structures (id: {name, fields, featureCount})
 */
const extraction = async ({ log }: SpreadsheetProcessingContext, tmpFile : string) => {
  await log.step('Récupération de la structure des données')

  // Display sheets
  const workbook = new Excel.Workbook()
  await workbook.xlsx.readFile(tmpFile)

  const sheetsList: SheetsList = {}

  for (const sheet of workbook.worksheets) {
    if (sheet.columnCount <= 0) {
      await log.warning(`Feuille ${sheet.id} - ${sheet.name} - Pas d'attributs, INUTILISABLE`)
    } else {
      await log.info(`Feuille ${sheet.id} - ${sheet.name} - ${sheet.actualRowCount - 1} lignes`)
      sheetsList[sheet.id] = { name: sheet.name, featureCount: sheet.actualRowCount - 1 }
    }
  }

  return sheetsList
}

/**
 * Allows you to create the requested sheet datasets
 * @param processingConfig  Processing configuration, obtained from the form data (processing-config-schema.json)
 * @param axios             Server for API requests
 * @param tmpDir            Directory where to download temporary files
 * @param log               Log system that is displayed on the user interface
 * @param sheetsList   Dictionary containing the structure of the file's sheets (id: {name, fields, featureCount})
 * @param tmpFile           Full path of the file to be processed
 * @returns   A list of objects associating sheets and datasets, or nothing at all to stop the program
 */
const createDatasets = async ({ processingConfig: rawConfig, axios, tmpDir, log } : SpreadsheetProcessingContext, sheetsList: SheetsList, tmpFile: string) => {
  // Narrow the union type to the create-mode branch (caller guarantees datasetMode === 'create').
  const processingConfig = rawConfig as CreateDatasets & Parameters
  await log.step('Construction des jeux de données')

  const sheetsTab: SheetsTab[] = []

  for (let sheet of processingConfig.sheets as SheetsTab[]) {
    if (sheet.add || processingConfig.addAllSheets) {
      sheet = {
        ...sheet,
        title: sheet.titleEditable ?? (sheet.name ?? 'untitled'),
        titleReadOnly: sheet.titleEditable ?? (sheet.name ?? 'untitled')
      }
      sheetsTab.push(sheet)
    }
  }

  // If there are no sheets to extract, we stop here to simplify the display of logs on the interface.
  if (sheetsTab.length <= 0) {
    await log.warning('Pas de feuilles renseignées')
    return
  }

  const sheetsTabCreate: SheetsTab[] = []
  const updateConfig = []

  // SECURITY (normally not necessary) : Checking the availability of the sheets (in the event that the download URL has been changed accidentally)
  for (const sheet of sheetsTab) {
    const idSheet = sheet.nb
    const nameSheet = sheet.name
    if (!(idSheet in sheetsList && sheetsList[idSheet].name === nameSheet)) {
      await log.warning(`La feuille ${idSheet} - ${nameSheet} n'est pas présente dans les couches disponibles`)
    } else {
      sheetsTabCreate.push(sheet)
    }
  }

  await log.info(`Extraction des feuilles ${sheetsTabCreate.map(sheet => (`${sheet.nb} - ${sheet.name}`)).join(', ')}`)

  for (const sheet of sheetsTabCreate) {
    if (shouldBeStopped) return

    const idSheet = sheet.nb

    await log.info(`Création du jeu de données pour la feuille ${idSheet} - ${sheetsList[idSheet].name}`)

    const tmpFileCSV = await createTmpFile(tmpDir, tmpFile, sheetsList[idSheet].name, log, () => shouldBeStopped)
    if (!tmpFileCSV) return

    const formData = new FormData()
    formData.append('title', sheet.title)
    formData.append('origin', processingConfig.url)
    formData.append('file', await fs.createReadStream(tmpFileCSV), { filename: path.parse(tmpFileCSV).base })
    const getLength = util.promisify(formData.getLength.bind(formData))
    const contentLength = await getLength()
    await log.info(`Chargement de ${formatBytes(contentLength)}`)

    if (shouldBeStopped) return

    const dataset = (await axios({
      method: 'post',
      url: 'api/v1/datasets',
      data: formData,
      maxContentLength: Infinity,
      maxBodyLength: Infinity,
      headers: { ...formData.getHeaders(), 'content-length': contentLength }
    })).data
    await log.info(`   Jeu de données créé, id="${dataset.id}", titre="${dataset.title}"`)

    const datasetObject = { id: dataset.id, href: dataset.href, title: dataset.title }
    const updateObject = { dataset: datasetObject, sheet: { nb: sheet.nb, name: sheet.name } }
    updateConfig.push(updateObject)

    await log.info('')
  }

  return { sheetsTab: processingConfig.sheets, updateConfig }
}

/**
 * Allows updating a dataset, either by force (schema reset) or by non-force (data replacement).
 * @param processingConfig  Processing configuration, obtained from the form data (processing-config-schema.json)
 * @param axios             Server for API requests
 * @param tmpDir            Directory where to download temporary files
 * @param log               Log system that is displayed on the user interface
 * @param sheetsList   Dictionary containing the structure of the file's sheets (id: {name, fields, featureCount})
 * @param tmpFile           Full path of the file to be processed
 * @returns   Returns nothing, used to stop the program
 */
const updateDatasets = async ({ processingConfig: rawConfig, axios, tmpDir, log } : SpreadsheetProcessingContext, sheetsList: SheetsList, tmpFile: string) => {
  // Narrow the union type to the update-mode branch (caller guarantees datasetMode === 'update').
  const processingConfig = rawConfig as UpdateDatasets
  await log.step('Mise à jour des jeux de données')

  // If there are no updates to extract, we stop here to simplify the display of logs on the interface.
  if (!processingConfig.datasets || processingConfig.datasets.length <= 0) {
    await log.warning('Pas de mises à jour renseignées')
    return
  }

  // ---------------------------------
  // SECURITY (normally not necessary): we verify that we have a file dataset
  // ---------------------------------

  const datasetsUpdate = []
  // Checking the availability of the sheets and the datasets
  for (const update of processingConfig.datasets) {
    if (!update.dataset.id || !update.dataset.title) {
      await log.warning('Le jeu de données est incomplet (id ou titre manquant)')
      await log.info('')
      continue
    }

    try {
      await axios.get(`api/v1/datasets/${update.dataset.id}`)
    } catch (err : any) {
      if (err.response?.status === 404) {
        await log.warning('Le jeu de données n\'existe pas. Il a peut-être été supprimé.')
      } else {
        await log.warning('L\'identification du jeu de données a échoué.')
      }
      await log.info('')
      continue
    }

    // SECURITY (normally not necessary) : Checking the availability of the sheets
    if (!(update.sheet.nb! in sheetsList && update.sheet.name! === sheetsList[update.sheet.nb!].name)) {
      await log.warning(`La feuille ${update.sheet.nb} - ${update.sheet.name} n'est pas présente dans les feuilles disponibles`)
      await log.info('')
      continue
    }

    datasetsUpdate.push(update)
  }

  // We process each dataset to be updated
  for (const update of datasetsUpdate) {
    if (shouldBeStopped) return

    const dataset = update.dataset
    const idSheet = update.sheet.nb!
    const sheetName = update.sheet.name!
    const formData = new FormData()

    await log.info(`Mise à jour du jeu ${dataset.title} avec la feuille ${idSheet} - ${sheetName}`)

    if (shouldBeStopped) return

    if (update.forceUpdate) await log.info('Mise à jour forcée du schéma')

    // Data update
    const tmpFileCSV = await createTmpFile(tmpDir, tmpFile, sheetsList[idSheet].name, log, () => shouldBeStopped)
    if (!tmpFileCSV) return

    formData.append('file', await fs.createReadStream(tmpFileCSV), { filename: path.parse(tmpFileCSV).base })
    formData.append('origin', processingConfig.url)
    const getLength = util.promisify(formData.getLength.bind(formData))
    const contentLength = await getLength()
    await log.info(`Chargement de ${formatBytes(contentLength)}`)

    if (shouldBeStopped) return

    await axios({
      method: 'post',
      url: `api/v1/datasets/${dataset.id}${update.forceUpdate ? '' : '?draft=true'}`,
      data: formData,
      maxContentLength: Infinity,
      maxBodyLength: Infinity,
      headers: { ...formData.getHeaders(), 'content-length': contentLength }
    })

    await log.info('')
  }
}
