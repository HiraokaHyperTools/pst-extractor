import { openPstFile } from '../openPstFile.js';
import { PSTFile } from '../PSTFile.class.js';
import { PSTFolder } from '../PSTFolder.class.js';
import { PSTMessage } from '../PSTMessage.class.js';
import { PSTRecipient } from '../PSTRecipient.class.js';
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
//  |  |  |  |  |-  sender: Lee  Dennis (Dennis.Lee@ENRON.com)
//  |  |  |  |  |-  recipient: Lindberg  Lorraine (Lorraine.Lindberg@ENRON.com)
//  |  |  |  |  |-  recipient: Watson  Kimberly (Kimberly.Watson@ENRON.com)
//  |  |  |  |  |-  recipient: Lee  Dennis (Dennis.Lee@ENRON.com)
//  |  |  |  |  |- Email: 2097220 - I/B Link Capacity for November and December 2001
//  |  |  |  |  |-  sender: Reames Julie (JReames@br-inc.com)
//  |  |  |  |  |-  recipient: Michelle Lokay (E-mail) (michelle.lokay@enron.com)

describe('PSTRecipient tests', () => {
  it('should have email messages', async () => {
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

    let msg: PSTMessage = (await comGroupFolder.getEmail(0))
    expect(msg.messageClass).toEqual('IPM.Note')
    expect(msg.subject).toEqual("New OBA's")
    expect(msg.senderName).toEqual('Lee  Dennis')
    expect(msg.senderEmailAddress).toEqual('Dennis.Lee@ENRON.com')
    expect(msg.displayTo).toEqual('Lindberg  Lorraine; Watson  Kimberly')

    let recipient: PSTRecipient = (await msg.getRecipient(0))
    expect(recipient).toBeTruthy()
    if (recipient) {
      // Log.debug1(JSON.stringify(recipient, null, 2));
      expect(recipient.displayName).toEqual('Lindberg  Lorraine')
      expect(recipient.smtpAddress).toEqual('Lorraine.Lindberg@ENRON.com')
    }

    recipient = (await msg.getRecipient(1))
    expect(recipient).toBeTruthy()
    if (recipient) {
      expect(recipient.displayName).toEqual('Watson  Kimberly')
      expect(recipient.smtpAddress).toEqual('Kimberly.Watson@ENRON.com')
    }

    recipient = (await msg.getRecipient(2))
    expect(recipient).toBeTruthy()
    if (recipient) {
      expect(recipient.displayName).toEqual('Lee  Dennis')
      expect(recipient.smtpAddress).toEqual('Dennis.Lee@ENRON.com')
    }

    msg = (await comGroupFolder.getEmail(1))
    expect(msg.messageClass).toEqual('IPM.Note')
    expect(msg.subject).toEqual(
      'I/B Link Capacity for November and December 2001'
    )
    expect(msg.sentRepresentingEmailAddress).toEqual('JReames@br-inc.com')
    expect(msg.displayTo).toEqual('Michelle Lokay (E-mail)')

    recipient = (await msg.getRecipient(0))
    expect(recipient).toBeTruthy()
    if (recipient) {
      // Log.debug1(JSON.stringify(recipient, null, 2));
      expect(recipient.displayName).toEqual('Michelle Lokay (E-mail)')
      expect(recipient.smtpAddress).toEqual('michelle.lokay@enron.com')
    }
  })
})
