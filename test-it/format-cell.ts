import { strict as assert } from 'node:assert'
import { it, describe } from 'node:test'
import { formatCell } from '../lib/tmp-file.ts'

describe('CSV cell formatting', () => {
  it('should flatten exceljs cell values', () => {
    assert.equal(formatCell(null), '')
    assert.equal(formatCell(12.5), '12.5')
    assert.equal(formatCell(new Date('2024-03-05T00:00:00Z')), '2024-03-05')
    assert.equal(formatCell(new Date('2024-03-05T10:30:00Z')), '2024-03-05T10:30:00.000Z')
    assert.equal(formatCell({ formula: 'A1*2', result: 4 }), '4')
    assert.equal(formatCell({ richText: [{ text: 'ab' }, { text: 'cd' }] }), 'abcd')
    assert.equal(formatCell({ text: 'site', hyperlink: 'https://a.b' }), 'site')
    assert.equal(formatCell({ error: '#N/A' }), '#N/A')
  })

  it('should quote values per RFC 4180', () => {
    assert.equal(formatCell('a,b'), '"a,b"')
    assert.equal(formatCell('say "hi"'), '"say ""hi"""')
    assert.equal(formatCell('a\r\nb'), '"a\r\nb"')
  })
})
