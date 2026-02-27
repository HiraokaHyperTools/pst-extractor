import { openPstFile } from '../openPstFile.js';
import { PSTFile } from '../PSTFile.class.js';
import { PSTFolder } from '../PSTFolder.class.js';
import { PSTMessage } from '../PSTMessage.class.js';
import { resolve } from 'path';
let pstFile: PSTFile

beforeAll(async () => {
  pstFile = await openPstFile(resolve('./src/__tests__/testdata/enron.pst'))
})

afterAll(async () => {
  await pstFile.close()
})

// get these emails
// Personal folders
//  |- Top of Personal Folders
//  |  |- Deleted Items
//  |  |- lokay-m
//  |  |  |- MLOKAY (Non-Privileged)
//  |  |  |  |- TW-Commercial Group
//  |  |  |  |  |- Email: 2097188 - New OBA's
//  |  |  |  |  |- Email: 2097220 - I/B Link Capacity for November and December 2001

describe('PSTObject tests', () => {
  it('should have basic attributes', async () => {
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
    childFolders = (await folder.getSubFolders())
    folder = childFolders[0]
    expect(folder.displayName).toEqual('MLOKAY (Non-Privileged)')
    childFolders = (await folder.getSubFolders())
    expect(childFolders[0].displayName).toEqual('TW-Commercial Group')
    const comGroupFolder = childFolders[0]
    // Log.debug1(JSON.stringify(comGroupFolder, null, 2));

    const msg: PSTMessage = (await comGroupFolder.getEmail(0))
    // Log.debug1(JSON.stringify(msg, null, 2));
    expect(msg.messageClass).toEqual('IPM.Note')
    expect(msg.messageSize.toNumber()).toEqual(653764)
  })
})
