import { openPstFile } from '../openPstFile.js';
import { PSTFile } from '../PSTFile.class.js';
import { resolve } from 'path';
let pstFile: PSTFile

beforeAll(async () => {
  pstFile = await openPstFile(resolve('./src/__tests__/testdata/enron.pst'))
})

afterAll(async () => {
  await pstFile.close()
})

describe('PSTfile tests', () => {
  it('should open the file', async () => {
    expect((await pstFile.getMessageStore()).displayName).toEqual('Personal folders')
    expect((await pstFile.getRootFolder())).toBeTruthy()
  })
})
