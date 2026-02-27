import { PSTFile } from '../PSTFile.class.js';
import { PSTMessage } from '../PSTMessage.class.js';
import { PSTAttachment } from '../PSTAttachment.class.js';
import { openPstFile } from '../openPstFile.js';
import { resolve } from 'path';
let pstFile: PSTFile

beforeAll(async () => {
  pstFile = await openPstFile(
    resolve('./src/__tests__/testdata/alpha-beta-gamma-delta.pst')
  )
})

afterAll(async () => {
  await pstFile.close()
})

describe('PSTAttachment tests', () => {
  it('should have a message with an triple embedded messages', async () => {
    // get to IPM_SUBTREE folder
    const folder = (await (await pstFile.getRootFolder()).getSubFolder(0))

    const alphaMsg: PSTMessage = (await folder.getEmail(0))
    expect(alphaMsg.messageClass).toEqual('IPM.Note')
    expect(alphaMsg.hasAttachments).toBeTruthy()
    expect((await alphaMsg.getNumberOfAttachments())).toEqual(2);
    const alphaAttachment1: PSTAttachment = (await alphaMsg.getAttachment(1))
    expect(alphaAttachment1.longFilename).toEqual('')
    expect(alphaAttachment1.size).toEqual(25634)
    expect(alphaAttachment1.modificationTime).toEqual(null)
    expect(alphaAttachment1.filename).toEqual("")
    expect(alphaAttachment1.attachMethod).toEqual(5)
    expect(alphaAttachment1.attachNum).toEqual(0)
    expect(alphaAttachment1.renderingPosition).toEqual(-1)
    expect(alphaAttachment1.mimeSequence).toEqual(0)
    expect(alphaAttachment1.pathname).toEqual('')
    expect(alphaAttachment1.longPathname).toEqual('')
    expect(alphaAttachment1.mimeTag).toEqual('')
    expect(alphaAttachment1.contentId).toEqual('')
    expect(alphaAttachment1.isAttachmentInvisibleInHtml).toEqual(false)
    expect(alphaAttachment1.isAttachmentInvisibleInRTF).toEqual(false)
    expect(alphaAttachment1.filesize).toEqual(22072);

    // The valid embeddedPSTMessage should be returned even if we request more than once.
    // We use `for` because we need async/await-able scope.
    for (let _ of [1, 2]) {
      const betaMsg = await alphaAttachment1.getEmbeddedPSTMessage();
      if (!betaMsg) {
        throw new Error("betaMsg must be available");
      }
      expect(betaMsg.hasAttachments).toBeTruthy()
      expect((await betaMsg.getNumberOfAttachments())).toEqual(2)

      const betaAttachment1: PSTAttachment = (await betaMsg.getAttachment(1));
      for (let _ of [1, 2]) {
        const gammaMsg = await betaAttachment1.getEmbeddedPSTMessage();
        if (!gammaMsg) {
          throw new Error("gammaMsg must be available");
        }
        expect(gammaMsg.hasAttachments).toBeTruthy()
        expect((await gammaMsg.getNumberOfAttachments())).toEqual(2)

        const gammaAttachment1: PSTAttachment = (await gammaMsg.getAttachment(1));
        for (let _ of [1, 2]) {
          const deltaMsg = await gammaAttachment1.getEmbeddedPSTMessage();
          if (!deltaMsg) {
            throw new Error("deltaMsg must be available");
          }
          expect(deltaMsg.hasAttachments).toBeTruthy()
          expect((await deltaMsg.getNumberOfAttachments())).toEqual(1)
        }
      }
    }
  })
})
