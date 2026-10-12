import { strict as assert } from 'node:assert'
import { it, describe, before, after } from 'node:test'
import http from 'node:http'
import type { AddressInfo } from 'node:net'
import os from 'node:os'
import path from 'node:path'
import fs from 'fs-extra'
import Excel from 'exceljs'
import testUtils from '@data-fair/lib-processing-dev/tests-utils.js'
import * as xlsxPlugin from '../index.ts'

import processingConfigSchema from '../processing-config-schema.json' with { type: 'json' }

/**
 * Fully local: a single HTTP server both serves the source xlsx file and stands in for data-fair,
 * recording every dataset upload so the produced CSV can be checked.
 */
describe('Spreadsheet processing', () => {
  let server: http.Server
  let baseUrl: string
  let tmpDir: string
  let xlsx: Buffer
  const uploads: { url: string, body: string }[] = []

  before(async () => {
    const workbook = new Excel.Workbook()
    workbook.addWorksheet('people').addRows([['name', 'age'], ['Alice', 30], ['Bob, Jr', 25]])
    workbook.addWorksheet('cities').addRows([['city'], ['Paris']])
    xlsx = Buffer.from(await workbook.xlsx.writeBuffer())

    server = http.createServer(async (req, res) => {
      if (req.method === 'GET' && req.url === '/sheets.xlsx') return res.end(xlsx)
      if (req.method === 'GET' && req.url?.startsWith('/api/v1/datasets/')) return res.end('{}')
      if (req.method === 'POST' && req.url?.startsWith('/api/v1/datasets')) {
        let body = ''
        for await (const chunk of req) body += chunk
        uploads.push({ url: req.url, body })
        res.setHeader('content-type', 'application/json')
        return res.end(JSON.stringify({ id: 'ds1', href: `${baseUrl}/api/v1/datasets/ds1`, title: 'people' }))
      }
      res.statusCode = 404
      res.end()
    })
    await new Promise<void>(resolve => server.listen(0, '127.0.0.1', resolve))
    baseUrl = `http://127.0.0.1:${(server.address() as AddressInfo).port}`
    tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), 'processing-spreadsheet-'))
  })

  after(async () => {
    server.close()
    await fs.remove(tmpDir)
  })

  const context = (processingConfig: any) => testUtils.context({ processingConfig, tmpDir }, { dataFairUrl: baseUrl, dataFairAPIKey: 'test' }, false)

  it('should expose a processing config schema for users', async () => {
    assert.equal(processingConfigSchema.type, 'object')
  })

  it('should list the sheets of a xlsx file', async () => {
    const ctx = context({ datasetMode: 'list', url: `${baseUrl}/sheets.xlsx` })
    await xlsxPlugin.run(ctx)
    assert.equal(ctx.processingConfig.datasetMode, 'create')
    assert.deepEqual(ctx.processingConfig.sheets.map((s: any) => [s.nb, s.name, s.lines]), [[1, 'people', 2], [2, 'cities', 1]])
  })

  it('should create a dataset from a sheet then update it', async () => {
    const ctx = context({
      datasetMode: 'create',
      url: `${baseUrl}/sheets.xlsx`,
      sheets: [{ add: true, nb: 1, name: 'people', lines: 2, titleEditable: 'people' }]
    })
    await xlsxPlugin.run(ctx)
    assert.equal(ctx.processingConfig.datasetMode, 'update')
    assert.equal(ctx.processingConfig.datasets[0].dataset.id, 'ds1')
    assert.equal(uploads.length, 1)
    assert.equal(uploads[0].url, '/api/v1/datasets')
    assert.ok(uploads[0].body.includes('name,age\nAlice,30\n"Bob, Jr",25\n'), uploads[0].body)

    await xlsxPlugin.run(ctx)
    assert.equal(uploads.length, 2)
    assert.equal(uploads[1].url, '/api/v1/datasets/ds1?draft=true')
  })
})
