import { pipeline } from 'node:stream/promises'
import type { AxiosRequestConfig } from 'axios'

import fs from 'fs-extra'
import path from 'path'

import type { SpreadsheetProcessingContext } from './context.ts'

/**
 * Allows you to download a file using the HTTP protocol
 * @param processingConfig  Processing configuration, obtained from the form data (processing-config-schema.json)
 * @param tmpFile           Name of the temporary file to transfer the downloaded file
 * @param axios             Server for API requests
 * @returns Name of extracted file
 */
export const fetchHTTP = async (processingConfig: SpreadsheetProcessingContext['processingConfig'], tmpFile: string, axios: SpreadsheetProcessingContext['axios']) => {
  const opts: AxiosRequestConfig = { responseType: 'stream', maxRedirects: 4 }

  // File retrieval and download
  let res
  try {
    res = await axios.get(processingConfig.url.trim(), opts)
  } catch (err: any) {
    if (err.response?.status === 404) throw new Error(`File not found: ${processingConfig.url.trim()}`)
    throw err
  }
  await pipeline(res.data, fs.createWriteStream(tmpFile))

  // Retrieving the file name
  if (processingConfig.filename) return processingConfig.filename
  if (res.headers['content-disposition'] && res.headers['content-disposition'].includes('filename=')) {
    if (res.headers['content-disposition'].match(/filename="?(.*?)"?;/)) return res.headers['content-disposition'].match(/filename="?(.*?)"?;/)[1]
    if (res.headers['content-disposition'].match(/filename="(.*)"/)) return res.headers['content-disposition'].match(/filename="(.*)"/)[1]
    if (res.headers['content-disposition'].match(/filename=(.*)/)) return res.headers['content-disposition'].match(/filename=(.*)/)[1]
  }
  if (res.request && res.request.res && res.request.res.responseUrl) return decodeURIComponent(path.parse(res.request.res.responseUrl).base)
}
