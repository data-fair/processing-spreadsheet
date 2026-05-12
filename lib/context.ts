import type { ProcessingContext } from '@data-fair/lib-common-types/processings.js'
import type { ProcessingConfig } from '#types/processingConfig/index.ts'

export type SpreadsheetProcessingContext = ProcessingContext<ProcessingConfig>

export type SheetsTab = {
  add: boolean,
  nb: number,
  name: string,
  lines: number,
  title?: string,
  titleEditable: string,
  titleReadOnly: string
}
