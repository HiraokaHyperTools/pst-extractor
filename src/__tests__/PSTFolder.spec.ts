import { openPstFile } from '../openPstFile'
import { PSTFile } from '../PSTFile.class'
import { PSTFolder } from '../PSTFolder.class'
const resolve = require('path').resolve
let pstFile: PSTFile

beforeAll(async () => {
  pstFile = await openPstFile(resolve('./src/__tests__/testdata/enron.pst'))
})

afterAll(async () => {
  await pstFile.close()
})

describe('PSTFolder tests', () => {
  it('should have a root folder', async () => {
    const folder: PSTFolder = (await pstFile.getRootFolder())
    expect(folder).toBeTruthy()
    expect(folder.displayName).toBe("")
    expect((await folder.getSubFolder(0)).displayName).toBe("Top of Personal Folders")
    expect((await folder.getSubFolder(1)).displayName).toBe("Search Root")
    expect((await folder.getSubFolderCount())).toEqual(2)
    expect(folder.hasSubfolders).toBeTruthy()
  })

  // folder structure should look like:
  // Personal folders
  //  |- Top of Personal Folders
  //  |  |- Deleted Items
  //  |  |- lokay-m
  //  |  |  |- MLOKAY (Non-Privileged)
  //  |  |  |  |- TW-Commercial Group
  //  |  |  |  |- Systems
  //  |  |  |  |- Sent Items
  //  |  |  |  |- Personal
  //  |- Search Root
  //  |- SPAM Search Folder 2

  it('root folder should have sub folders', async () => {
    let childFolders: PSTFolder[] = (await (await pstFile.getRootFolder()).getSubFolders())
    expect(childFolders.length).toEqual(2)
    let folder = childFolders[0]
    expect((await folder.getSubFolderCount())).toEqual(2)
    expect(folder.displayName).toEqual('Top of Personal Folders')
    childFolders = (await folder.getSubFolders())
    folder = childFolders[0]
    expect(folder.displayName).toEqual('Deleted Items')
    folder = childFolders[1]
    expect(folder.displayName).toEqual('lokay-m')
    // Log.debug1(JSON.stringify(folder, null, 2));
    childFolders = (await folder.getSubFolders())
    folder = childFolders[0]
    expect(folder.displayName).toEqual('MLOKAY (Non-Privileged)')
    childFolders = (await folder.getSubFolders())
    expect(childFolders[0].displayName).toEqual('TW-Commercial Group')
    expect(childFolders[1].displayName).toEqual('Systems')
    expect(childFolders[2].displayName).toEqual('Sent Items')
    expect(childFolders[3].displayName).toEqual('Personal')
    expect((await folder.getSubFolderCount())).toEqual(4)
    expect((await folder.getEmailCount())).toEqual(1)
    expect(folder.folderType).toEqual(0)
    expect(folder.contentCount).toEqual(1)
    expect(folder.unreadCount).toEqual(0)
    expect(folder.containerFlags).toEqual(0)
    expect(folder.containerClass).toEqual('IPF.Note')
    expect(folder.hasSubfolders).toEqual(true)
    // Log.debug1(JSON.stringify(folder, null, 2));

    await folder.getEmail(0)
    expect((async () => { await folder.getEmail(1) })()).rejects.toThrow(RangeError)
    expect((async () => { await folder.getEmail(100) })()).rejects.toThrow(RangeError)
  })
})
