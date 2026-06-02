import config from '#config'
import { strict as assert } from 'node:assert'
import { it, describe } from 'node:test'
import testUtils from '@data-fair/lib-processing-dev/tests-utils.js'
import * as xlsxPlugin from '../index.ts'

import processingConfigSchema from '../processing-config-schema.json' with { type: 'json' }

/**
 * Used to test the list of sofas in a file and the correct creation of datasets from a file.
 * We do not test the update because we cannot retrieve the necessary information for the test.
*/
describe('Geopackage processing', () => {
  // Each plugin should expose a processing config schema

  it('should expose a processing config schema for users', async () => {
    assert.equal(processingConfigSchema.type, 'object')
  })

  it('should display the layers of a xlsx file', async () => {
    const context = testUtils.context({
      pluginConfig: {},
      processingConfig: {
        datasetMode: 'list',
        url: 'https://www.data.gouv.fr/api/1/datasets/r/aa7a0f1c-89e3-4d40-af94-6f226202ada3',
      },
      tmpDir: 'test-data.test/'
    }, config, false)

    await xlsxPlugin.run(context)
  })

  it('should run a task with a xlsx file to create an file dataset', async function () {
    const context = testUtils.context({
      pluginConfig: {},
      processingConfig: {
        datasetMode: 'create',
        url: 'https://www.data.gouv.fr/api/1/datasets/r/aa7a0f1c-89e3-4d40-af94-6f226202ada3',
        sheets: [
          {
            add: true,
            nb: 3,
            name: 'fra',
            lines: 539,
            titleEditable: 'test-xlsx-3'
          }
        ]
      },
      tmpDir: 'test-data.test/'
    }, config, false)

    await xlsxPlugin.run(context)
    assert.equal(context.processingConfig.datasetMode, 'update')
  })
})
