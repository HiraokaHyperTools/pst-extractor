import { PSTFile } from '../PSTFile.class.js';
import { PSTFolder } from '../PSTFolder.class.js';
import { PSTActivity } from '../PSTActivity.class.js';
import { openPstFile } from '../index.js';
import { resolve } from 'path';
let pstFile: PSTFile
let folder: PSTFolder

beforeAll(async () => {
  pstFile = await openPstFile(
    resolve('./src/__tests__/testdata/mtnman1965@outlook.com.ost')
  )

  // get to Journal folder
  let childFolders: PSTFolder[] = (await (await pstFile.getRootFolder()).getSubFolders())
  folder = childFolders[1] // Root - Mailbox
  expect(folder.displayName).toBe("Root - Mailbox");
  childFolders = await folder.getSubFolders()
  folder = childFolders[4] // IPM_SUBTREE
  expect(folder.displayName).toBe("IPM_SUBTREE");
  childFolders = await folder.getSubFolders()
  folder = childFolders[15] // Journal
  expect(folder.displayName).toBe("Journal");
})

afterAll(async () => {
  await pstFile?.close()
})

describe('PSTActivity tests', () => {
  it('should have a Journal folder', () => {
    expect(folder.displayName).toEqual('Journal')
  })

  it('root folder should have a journal entry', async () => {
    const activity: PSTActivity = (await folder.getEmail(0)) as PSTActivity
    // console.log(JSON.stringify(activity, null, 2));
    expect(activity.messageClass).toEqual('IPM.Activity')
    expect(activity.subject).toEqual('called Ed')
    expect(activity.logTypeDesc).toEqual('Phone call')
    expect(activity.bodyPrefix).toContain('But no one was home')
    expect(activity.logStart).toEqual(new Date('2018-03-06T21:09:00.000Z'))
    expect(activity.logEnd).toEqual(new Date('2018-03-06T21:09:00.000Z'))
    expect(activity.importance).toEqual(1)
    expect(activity.logDuration).toEqual(0)
    expect(activity.logFlags).toEqual(0)
    expect(activity.isDocumentPrinted).toEqual(false)
    expect(activity.isDocumentSaved).toEqual(false)
    expect(activity.isDocumentRouted).toEqual(false)
    // expect(activity.isDocumentPosted).toEqual(false)
  })
})
