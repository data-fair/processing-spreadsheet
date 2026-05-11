import type { RunFunction } from '@data-fair/lib-common-types/processings.js'
import type { ProcessingConfig, CreateDatasets, UpdateDatasets, Parameters } from '#types/processingConfig/index.ts'
import { formatBytes } from '@data-fair/lib-utils/format/bytes.js'

import util from 'util'
import fs from 'fs-extra'
import path from 'path'
import Excel from 'exceljs'
import FormData from 'form-data'

import type { SpreadsheetProcessingContext } from './context.ts'
import { fetchHTTP } from './fetch.ts'
import { createTmpFile } from './tmp-file.ts'
import { runCommand } from './spawn-process.ts'

/**
 * Allows for a requested program shutdown to be scheduled.
 *
 * `stopSignal` is a promise that resolves the moment `stop()` is called. Long-running
 * waiters (typically `ws.waitForJournal` in phase 2) race against it so they can bail
 * out immediately instead of timing out after several minutes.
 */
let shouldBeStopped = false
let stopSignal: Promise<void> = new Promise(() => {})
let resolveStop: () => void = () => {}

export const stop: () => Promise<void> = async () => {
  shouldBeStopped = true
  resolveStop()
}

type SheetsList = Record<number, { name: string, featureCount: number }>

type PendingFinalization = {
  promise: Promise<{ ok: true, journal: any } | { ok: false, error: Error }>
  datasetId: string
  datasetTitle: string
}

let nbFinalize = 0

/**
 * Starts listening for the `finalize-end` journal event of a dataset without blocking.
 *
 * `ws.waitForJournal` is invoked synchronously so the WebSocket subscription is set
 * up immediately after the upload — this keeps the race window between "event emitted
 * by the server" and "listener attached on the client" down to a single roundtrip,
 * matching the behaviour of the original sequential flow. The returned promise never
 * rejects: failures are converted into a warning log and an `{ ok: false }` result so
 * one bad finalization cannot abort `Promise.allSettled` over the whole batch. The
 * wait also races against the module-level `stopSignal`, so a `stop()` triggers an
 * immediate bail-out without waiting for the journal timeout.
 *
 * @param ws                    Data Fair WebSocket client used to receive journal events.
 * @param log                   Log system displayed in the user interface.
 * @param datasetId             Id of the dataset whose finalization should be awaited.
 * @param datasetTitle          Human-readable dataset title, used in log messages.
 * @param opts.successMessage   Message logged when the finalization succeeds.
 * @param opts.checkDraft       When true, a draft state on the journal triggers a schema-
 *                              incompatibility warning instead of the success message
 *                              (used by update flows).
 * @param progressInfo.name     Name of the corresponding task log
 * @param progressInfo.total    Total number of pending datasets
 * @returns A `PendingFinalization` whose `promise` settles once the event arrives,
 *          the run is stopped, the wait times out, or fails — never rejects.
 */
const trackFinalization = (
  ws: SpreadsheetProcessingContext['ws'],
  log: SpreadsheetProcessingContext['log'],
  datasetId: string,
  datasetTitle: string,
  opts: { successMessage: string, checkDraft?: boolean },
  progressInfo: { name: string, total: number }
): PendingFinalization => {
  const journalPromise = ws.waitForJournal(datasetId, 'finalize-end')
    .then(journal => ({ kind: 'event' as const, journal }))
  const stopPromise = stopSignal.then(() => ({ kind: 'stopped' as const }))

  const promise = Promise.race([journalPromise, stopPromise])
    .then(async (result) => {
      if (result.kind === 'stopped') {
        return { ok: false as const, error: new Error('stopped') }
      }
      const journal: any = result.journal
      if (opts.checkDraft && (journal.draft !== undefined || journal.draft)) {
        await log.warning(`Le schéma du jeu de données "${datasetTitle}" n'est pas compatible avec la couche . Le jeu est passé en mode brouillon, à vous de le valider ou non.`)
      } else {
        await log.info(`Le jeu de données "${datasetTitle}" ${opts.successMessage}`)
      }
      return { ok: true as const, journal }
    })
    .catch(async (error: Error) => {
      await log.warning(`Le jeu de données "${datasetTitle}" n'a pas pu être finalisé (${error.message}), vous pouvez relancer son traitement.`)
      return { ok: false as const, error }
    })
    .finally(async () => {
      nbFinalize += 1
      await log.progress(progressInfo.name, nbFinalize, progressInfo.total)
    })
  return { promise, datasetId, datasetTitle }
}

/**
 * Input function, allows data processing to begin
 * @param context Context of the request
 */
export const run: RunFunction<ProcessingConfig> = async (context) => {
  shouldBeStopped = false
  stopSignal = new Promise<void>(resolve => { resolveStop = resolve })

  try {
    // Retrieving the contextual elements necessary for processing
    const { processingConfig, patchConfig } = context
    const tmpFile = await download(context)

    if (shouldBeStopped) return
    if (!tmpFile) return
    const sheetsList = await extraction(context, tmpFile)

    if (shouldBeStopped) return
    if (!sheetsList) return

    if (processingConfig.datasetMode === 'create') {
      const updateConfig = await createDatasets(context, sheetsList, tmpFile)

      // The lib-common-types signature only allows `dataset` (singular), but the worker's
      // patchConfig is a generic Object.assign on the config — `datasets` is supported at runtime.
      if (updateConfig?.length) await patchConfig({ datasetMode: 'update', datasets: updateConfig } as any)
    } else if (processingConfig.datasetMode === 'update') {
      await updateDatasets(context, sheetsList, tmpFile)
    } else {
      await patchConfig({ datasetMode: 'create', dataset: { prefix: '' } })
    }
  } finally {
    // Settle the stop signal so any continuation chained on it (the `stopSignal.then(...)`
    // branches inside `trackFinalization`) is released. At this point all the relevant
    // `Promise.race` calls have already resolved via the journal branch, so this late
    // resolution is a no-op behaviour-wise — it only frees handlers.
    resolveStop()
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

  const url = new URL(processingConfig.url)
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
    await log.info('Le format n\'est pas pris en charge')
    throw new Error('Format non pris en charge')
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
 * @param ws                Data Fair's Websocket allows retrieving the dataset response.
 * @param sheetsList   Dictionary containing the structure of the file's sheets (id: {name, fields, featureCount})
 * @param tmpFile           Full path of the file to be processed
 * @returns   A list of objects associating sheets and datasets, or nothing at all to stop the program
 */
const createDatasets = async ({ processingConfig: rawConfig, axios, dir, log, ws } : SpreadsheetProcessingContext, sheetsList: SheetsList, tmpFile: string) => {
  // Narrow the union type to the create-mode branch (caller guarantees datasetMode === 'create').
  const processingConfig = rawConfig as CreateDatasets & Parameters
  await log.step('Construction des jeux de données')

  let idsSheets: number[] = []

  // If we want to add all the sheets, we add all the identifiers to the list.
  if (processingConfig.addAllSheets) {
    idsSheets = Object.keys(sheetsList).map(sheet => Number(sheet))
  } else {
    processingConfig.listIdsSheets = processingConfig.listIdsSheets ? processingConfig.listIdsSheets.replaceAll(' ', '') : ''

    const listParts = processingConfig.listIdsSheets.split(',')

    for (const part of listParts) {
      const idSheet = Number(part)

      if (idSheet && idSheet > 0) {
        idsSheets.push(idSheet)
      } else {
        const interval = part.split('-')

        if (interval.length === 2) {
          const start = Number(interval[0])
          const end = Number(interval[1])

          if (start && start > 0 && end && end >= start) {
            for (let id = start; id <= end; id++) {
              idsSheets.push(id)
            }
          }
        }
      }
    }
  }

  // If there are no sheets to extract, we stop here to simplify the display of logs on the interface.
  if (idsSheets.length <= 0) {
    await log.warning('Pas de feuilles renseignées')
    return
  }

  await log.info(`Extraction des feuilles ${idsSheets}`)

  const idsSheetsCreate = []
  const updateConfig = []
  const pendingFinalizations: PendingFinalization[] = []
  const progressName = 'En attente de la finalisation de la création des jeux de données'

  // Checking the availability of the sheets
  for (const idSheet of idsSheets) {
    if (!(idSheet in sheetsList)) {
      await log.warning(`La feuille ${idSheet} n'est pas présente dans les feuilles disponibles`)
    } else {
      idsSheetsCreate.push(idSheet)
    }
  }
  await log.info('')

  for (const idSheet of idsSheetsCreate) {
    if (shouldBeStopped) return

    await log.info(`Création du jeu de données pour la feuille ${idSheet} - ${sheetsList[idSheet].name}`)

    const tmpFileXLSX = await createTmpFile(dir, tmpFile, sheetsList[idSheet].name, log, () => shouldBeStopped)
    if (!tmpFileXLSX) return

    const formData = new FormData()
    formData.append('title', `${processingConfig.dataset.prefix} - ${sheetsList[idSheet].name}`)
    formData.append('file', await fs.createReadStream(tmpFileXLSX), { filename: path.parse(tmpFileXLSX).base })
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

    pendingFinalizations.push(trackFinalization(ws, log, dataset.id, dataset.title, { successMessage: 'a été finalisé' },
      { name: progressName, total: idsSheetsCreate.length }))

    const datasetObject = { id: dataset.id, href: dataset.href, title: dataset.title }
    const updateObject = { dataset: datasetObject, idSheet }
    updateConfig.push(updateObject)

    await log.info('')
  }

  if (pendingFinalizations.length > 0) {
    await log.step('Finalisation des jeux de données')

    await log.task(progressName)
    await log.progress(progressName, nbFinalize, idsSheetsCreate.length)

    await Promise.allSettled(pendingFinalizations.map(p => p.promise))
  }

  return updateConfig
}

/**
 * Allows updating a dataset, either by force (schema reset) or by non-force (data replacement).
 * @param processingConfig  Processing configuration, obtained from the form data (processing-config-schema.json)
 * @param axios             Server for API requests
 * @param tmpDir            Directory where to download temporary files
 * @param log               Log system that is displayed on the user interface
 * @param ws                Data Fair's Websocket allows retrieving the dataset response.
 * @param sheetsList   Dictionary containing the structure of the file's sheets (id: {name, fields, featureCount})
 * @param tmpFile           Full path of the file to be processed
 * @returns   Returns nothing, used to stop the program
 */
const updateDatasets = async ({ processingConfig: rawConfig, axios, dir, log, ws } : SpreadsheetProcessingContext, sheetsList: SheetsList, tmpFile: string) => {
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

  // We add size=10000 to ensure that all datasets are retrieved (12 by default)
  const datasets = (await axios.get<{ results: { id: string }[] }>('api/v1/datasets/?size=10000&file=true')).data.results
  const datasetsIds = new Set<string>(datasets.map(d => d.id))

  const datasetsUpdate = []
  // Checking the availability of the sheets and the datasets
  for (const update of processingConfig.datasets) {
    if (!update.dataset.id || !update.dataset.title) {
      await log.warning('Le jeu de données est incomplet (id ou titre manquant)')
      await log.info('')
      continue
    }

    // Check if the sheet is available
    if (!(update.idSheet in sheetsList)) {
      await log.warning(`La feuille ${update.idSheet} n'est pas présente dans les feuilles disponibles`)
      await log.info('')
      continue
    }

    // Check if the correct update operation can be performed, to avoid permission errors
    if (!(datasetsIds.has(update.dataset.id))) {
      await log.warning(`Le jeu de données ${update.dataset.title} n'est pas de type fichier`)
      await log.info('')
      continue
    }
    datasetsUpdate.push(update)
  }

  const pendingFinalizations: PendingFinalization[] = []
  const progressName = 'En attente de la finalisation de la mise à jour des jeux de données'

  // We process each dataset to be updated
  for (const update of datasetsUpdate) {
    if (shouldBeStopped) return

    const dataset = update.dataset
    const idSheet = update.idSheet
    const formData = new FormData()

    await log.info(`Mise à jour du jeu ${dataset.title} avec la feuille ${idSheet}`)

    if (shouldBeStopped) return

    if (update.forceUpdate) await log.info('Mise à jour forcée du schéma')

    // Data update
    const tmpFileXLSX = await createTmpFile(dir, tmpFile, sheetsList[idSheet].name, log, () => shouldBeStopped)
    if (!tmpFileXLSX) return

    formData.append('file', await fs.createReadStream(tmpFileXLSX), { filename: path.parse(tmpFileXLSX).base })
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

    // We are certain of the ID and title definitions with the previous check.
    pendingFinalizations.push(trackFinalization(ws, log, dataset.id!, dataset.title!, { successMessage: 'a été mis à jour', checkDraft: true },
      { name: progressName, total: datasetsUpdate.length }
    ))

    await log.info('')
  }

  if (pendingFinalizations.length > 0) {
    await log.step('Finalisation des mises à jour')

    await log.task(progressName)
    await log.progress(progressName, 0, datasetsUpdate.length)

    await Promise.allSettled(pendingFinalizations.map(p => p.promise))
  }
}
