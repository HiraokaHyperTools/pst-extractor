import { openPstFile } from '../openPstFile'
import { PSTFile } from '../PSTFile.class'
const resolve = require('path').resolve
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
