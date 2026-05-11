import fs from 'fs-extra'
import path from 'path'
import Excel from 'exceljs'

import type { SpreadsheetProcessingContext } from './context.ts'

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
  const tmpFileCSV = path.join(dir, `${sheetName}.csv`)

  if (await fs.pathExists(tmpFileCSV)) return tmpFileCSV

  await log.info('Création du fichier temporaire')
  if (isStopped()) return

  // Creating a stream for large files
  const workbookReader = new Excel.stream.xlsx.WorkbookReader(tmpFile, {})
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

      writeStream.write(line + '\n')
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
const formatCell = (cell: Excel.CellValue): string => {
  if (cell === null || cell === undefined) return ''
  const str = String(cell)

  // Escape if the value contains a comma, a quotation mark, or a line break
  if (str.includes(',') || str.includes('"') || str.includes('\n')) {
    return `"${str.replace(/"/g, '""')}"`
  }
  return str
}
