import { strict as assert } from 'node:assert'
import { describe, it, after } from 'node:test'
import fs from 'fs-extra'
import os from 'node:os'
import path from 'node:path'
import { safeLocalName, removeSymlinks } from '../lib/tmp-file.ts'

const tmpDirs: string[] = []
after(async () => {
  for (const dir of tmpDirs) await fs.remove(dir)
})

describe('tmp files', () => {
  it('builds local names without a directory part', () => {
    assert.equal(safeLocalName('Feuil1'), 'Feuil1')
    assert.equal(safeLocalName('../../other/run'), '.._.._other_run')
    assert.equal(path.dirname(path.join('/tmp/run', `${safeLocalName('..\\x/..')}.csv`)), '/tmp/run')
  })

  it('removes the symbolic links extracted from an archive', async () => {
    const dir = await fs.mkdtemp(path.join(os.tmpdir(), 'tmp-file-test-'))
    tmpDirs.push(dir)
    await fs.writeFile(path.join(dir, 'a.file'), 'a')
    await fs.symlink('/etc/hostname', path.join(dir, 'link.file'))
    await removeSymlinks(dir)
    assert.deepEqual(await fs.readdir(dir), ['a.file'])
  })
})
