import { PSTContact } from '../PSTContact.class.js';
import { PSTFile } from '../PSTFile.class.js';
import { PSTFolder } from '../PSTFolder.class.js';
import { PSTMessage } from '../PSTMessage.class.js';
import { PSTTask } from '../PSTTask.class.js';
import { PSTAttachment } from '../PSTAttachment.class.js';
import { openPstFile } from '../index.js';
import { resolve } from 'path';
let pstFile: PSTFile
let subtreeFolder: PSTFolder

beforeAll(async () => {
  pstFile = await openPstFile(
    resolve('./src/__tests__/testdata/mtnman1965@outlook.com.ost')
  )

  // get to IPM_SUBTREE folder
  let childFolders: PSTFolder[] = await (await (await pstFile.getRootFolder()).getSubFolders())
  subtreeFolder = childFolders[1] // Root - Mailbox
  expect(subtreeFolder.displayName).toBe("Root - Mailbox");
  childFolders = await subtreeFolder.getSubFolders()
  subtreeFolder = childFolders[4] // IPM_SUBTREE
  expect(subtreeFolder.displayName).toBe("IPM_SUBTREE");
})

afterAll(async () => {
  await pstFile.close()
})

describe('PSTAttachment tests', () => {
  it('should have a contact with an attachment', async () => {
    const folder = await subtreeFolder.getSubFolder(10) // Contacts
    expect(folder.displayName).toBe("Contacts");
    const contact: PSTContact = await folder.getEmail(0) as PSTContact
    expect(contact.messageClass).toEqual('IPM.Contact')
    expect(contact.hasAttachments).toBeTruthy()

    // first attachment is contact picture
    let attachment: PSTAttachment = await (await contact.getAttachment(0))
    expect(attachment.size).toEqual(14055)
    expect(attachment.longFilename).toEqual('ContactPicture.jpg')
    expect(attachment.creationTime).toEqual(
      new Date('2018-03-07T16:50:00.000Z')
    )

    // second attachment is never gonna give you up
    attachment = await (await contact.getAttachment(1))
    expect(attachment.size).toEqual(8447)
    expect(attachment.longFilename).toEqual('rickroll.jpg')
    expect(attachment.creationTime).toEqual(
      new Date('2018-03-07T16:49:32.964Z')
    )
  })

  it('should have a task with an attachment', async () => {
    const folder = await subtreeFolder.getSubFolder(17) // Tasks
    expect(folder.displayName).toBe('Tasks');
    const task: PSTTask = (await folder.getEmail(0)) as PSTTask
    expect(task.messageClass).toEqual('IPM.Task')
    expect(task.hasAttachments).toBeTruthy()
    const attachment: PSTAttachment = (await task.getAttachment(0))
    expect(attachment.size).toEqual(8447)
    expect(attachment.longFilename).toEqual('rickroll.jpg')
    expect(attachment.creationTime).toEqual(
      new Date('2018-03-07T16:49:32.964Z')
    )
    expect(attachment.modificationTime).toEqual(
      new Date('2018-03-07T16:49:32.959Z')
    )
    expect(attachment.filename).toEqual('rickroll.jpg')
    expect(attachment.attachMethod).toEqual(1)
    expect(attachment.attachNum).toEqual(0)
    expect(attachment.renderingPosition).toEqual(60)
    expect(attachment.mimeSequence).toEqual(0)
    expect(attachment.pathname).toEqual('')
    expect(attachment.longPathname).toEqual('')
    expect(attachment.mimeTag).toEqual('')
    expect(attachment.contentId).toEqual('')
    expect(attachment.isAttachmentInvisibleInHtml).toEqual(false)
    expect(attachment.isAttachmentInvisibleInRTF).toEqual(false)
    expect(attachment.filesize).toEqual(4796)
    expect(await attachment.getEmbeddedPSTMessage()).toEqual(null)
  })

  it('should have email with word attachment', async () => {
    const folder = await subtreeFolder.getSubFolder(1) // Inbox
    let msg: PSTMessage = await folder.getEmail(0)
    msg = await folder.getEmail(1)
    msg = await folder.getEmail(2)

    // Email: 2110308 - word attachment
    expect(msg.hasAttachments).toBeTruthy()
    const attachment: PSTAttachment = (await msg.getAttachment(0))
    expect(attachment.size).toEqual(54044)
    expect(attachment.longFilename).toEqual('OBA_2760.doc')
    expect(attachment.creationTime).toEqual(
      new Date('2018-03-07T16:26:20.724Z')
    )
  })

  it('should have email with excel attachment', async () => {
    const folder = (await subtreeFolder.getSubFolder(1)) // Inbox
    let msg: PSTMessage = (await folder.getEmail(0))
    msg = (await folder.getEmail(1))
    msg = (await folder.getEmail(2))
    msg = (await folder.getEmail(3))

    // Email: 2110724 - excel attachment
    expect(msg.hasAttachments).toBeTruthy()
    const attachment: PSTAttachment = (await msg.getAttachment(0))
    expect(attachment.size).toEqual(31016)
    expect(attachment.longFilename).toEqual('RedRockA.xls')
    expect(attachment.creationTime).toEqual(
      new Date('2018-03-07T16:31:56.075Z')
    )
  })

  it('should have email with jpg attachment', async () => {
    const folder = (await subtreeFolder.getSubFolder(1)) // Inbox
    let msg: PSTMessage = (await folder.getEmail(0))
    msg = (await folder.getEmail(1))
    msg = (await folder.getEmail(2))
    msg = (await folder.getEmail(3))
    msg = (await folder.getEmail(4))

    // Email: 2111140 - never gonna give you up
    expect(msg.hasAttachments).toBeTruthy()
    const attachment: PSTAttachment = (await msg.getAttachment(0))
    expect(attachment.size).toEqual(5020)
    expect(attachment.longFilename).toEqual('rickroll.jpg')
    expect(attachment.creationTime).toEqual(
      new Date('2018-03-07T16:43:36.995Z')
    )
  })
})
