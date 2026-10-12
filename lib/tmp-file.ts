import { once } from 'node:events'
import fs from 'fs-extra'
import path from 'path'
import Excel from 'exceljs'

import type { SpreadsheetProcessingContext } from './context.ts'

/**
 * Name of a local file built from a name read in the downloaded file (sheet or layer name):
 * no directory part, so that it stays in the tmp directory
 */
export const safeLocalName = (name: string) => name.replace(/[/\\]/g, '_')

/**
 * unzip restores the symbolic links of an archive: reading or writing one of
 * them would reach a file of the worker outside of the tmp directory.
 */
export const removeSymlinks = async (dir: string): Promise<void> => {
  for (const f of await fs.readdir(dir)) {
    if ((await fs.lstat(path.join(dir, f))).isSymbolicLink()) await fs.remove(path.join(dir, f))
  }
}

type NamedWorksheetReader = Excel.stream.xlsx.WorksheetReader & { name: string, destroy: () => void }

/**
 * Allows you to create a temporary .xlsx file from a sheet in the data file, to be sent to create a file dataset.
 * @param dir         Directory where to store the file
 * @param tmpFile     Name of the temporary file containing the original data (multi-sheeted xlsx)
 * @param sheetName   Name of the sheet to be extracted
 * @param log         Log system that is displayed on the user interface
 * @param isStopped   Function allowing the program to stop if requested
 * @returns   Name of the temporary CSV file created to send
 */
export const createTmpFile = async (dir : string, tmpFile : string, sheetName : string, log: SpreadsheetProcessingContext['log'], isStopped: () => boolean) => {
  const tmpFileCSV = path.join(dir, `${safeLocalName(sheetName)}.csv`)

  if (await fs.pathExists(tmpFileCSV)) return tmpFileCSV

  await log.info('Création du fichier temporaire')
  if (isStopped()) return

  // Creating a stream for large files
  // styles are needed to recognise date cells, otherwise they come out as serial numbers
  const workbookReader = new Excel.stream.xlsx.WorkbookReader(tmpFile, { styles: 'cache' })
  const writeStream = fs.createWriteStream(tmpFileCSV, { encoding: 'utf-8' })

  // Retrieving the lines from the correct sheet
  for await (const worksheetReader of workbookReader) {
    // Try to fix a typing problem in ExcelJS
    const reader = worksheetReader as unknown as NamedWorksheetReader
    if (isStopped()) {
      writeStream.destroy()
      return
    }

    if (reader.name !== sheetName) {
      continue
    }

    for await (const row of worksheetReader) {
      if (isStopped()) {
        writeStream.destroy()
        return
      }

      const line = (row.values as Excel.CellValue[])
        .slice(1) // row.values is indexed starting from 1, 0 is null
        .map(cell => formatCell(cell))
        .join(',')

      if (!writeStream.write(line + '\n')) await once(writeStream, 'drain')
    }

    // We only want to build one file for one layer, so we stop the program once the task is completed.
    break
  }

  if (isStopped()) {
    writeStream.destroy()
    return
  }

  await new Promise<void>((resolve, reject) => {
    writeStream.end(resolve)
    writeStream.on('error', reject)
  })

  return tmpFileCSV
}

/**
 * Formats a cell value to be safely embedded in a CSV field.
 * Wraps the value in double quotes if it contains a comma, a double quote, or a line break,
 * and escapes any internal double quotes by doubling them (RFC 4180).
 *
 * @param cell  Cell value to format
 * @returns   CSV-safe string representation of the cell
 */
export const formatCell = (cell: Excel.CellValue): string => {
  if (cell === null || cell === undefined) return ''
  let str: string
  if (cell instanceof Date) {
    const iso = cell.toISOString()
    str = iso.endsWith('T00:00:00.000Z') ? iso.slice(0, 10) : iso
  } else if (typeof cell === 'object') {
    // formula, shared formula, rich text, hyperlink or error cells
    if ('result' in cell || 'formula' in cell || 'sharedFormula' in cell) return formatCell((cell as Excel.CellFormulaValue).result as Excel.CellValue)
    if ('richText' in cell) str = cell.richText.map(r => r.text).join('')
    else if ('text' in cell) return formatCell(cell.text as Excel.CellValue)
    else if ('error' in cell) str = cell.error
    else str = JSON.stringify(cell)
  } else {
    str = String(cell)
  }

  // Escape if the value contains a comma, a quotation mark, or a line break
  if (/[,"\r\n]/.test(str)) {
    return `"${str.replace(/"/g, '""')}"`
  }
  return str
}
