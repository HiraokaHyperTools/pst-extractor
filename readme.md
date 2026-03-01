# pst-extractor

[![npm](https://img.shields.io/npm/v/@hiraokahypertools/pst-extractor)](https://www.npmjs.com/package/@hiraokahypertools/pst-extractor)

Extract objects from MS Outlook/Exchange PST files

This is a private fork of https://github.com/epfromer/pst-extractor repository.

Links: [typedoc](https://hiraokahypertools.github.io/pst-extractor/typedoc/), [online demo](https://hiraokahypertools.github.io/pst-extractor-demo/)

## Install

```npm install --save @hiraokahypertools/pst-extractor```

or

```yarn add @hiraokahypertools/pst-extractor```

## Open .PST (.OST) file on Node.js

```ts
import { openPstFile } from '@hiraokahypertools/pst-extractor';

const pstFile = await openPstFile(
  'path/to/file.pst'
);
```

## Open .PST (.OST) file on web browser

```ts
import { openPst, type IPSTFile } from '@hiraokahypertools/pst-extractor';

async function openPstWithFile(file: File): Promise<IPSTFile> {
  return await openPst(
    {
      readFile: async (buffer: ArrayBuffer, offset: number, length: number, position: number): Promise<number> => {
        const blob = file.slice(position, position + length);
        const arrayBuffer = await blob.arrayBuffer();
        const srcArray = new Uint8Array(arrayBuffer);
        const destArray = new Uint8Array(buffer);
        destArray.set(srcArray, offset);
        return destArray.byteLength;
      },
      close: async (): Promise<void> => {

      }
    }
  );
}
```

## Traversal the entire .PST (.OST) file

```ts
import { Consts, PSTAttachment, type IPSTFolder } from '@hiraokahypertools/pst-extractor';

async function traverseFolder(folder: IPSTFolder) {
    console.log(folder.displayName);

    const messages = await folder.getEmails();
    for (const message of messages) {
        console.log(message.messageClass, message.displayName);

        const recipients = await message.getRecipients();
        for (const recipient of recipients) {
            if (recipient.recipientType == Consts.MAPI_TO) { }
            if (recipient.recipientType == Consts.MAPI_CC) { }
            if (recipient.recipientType == Consts.MAPI_BCC) { }

            console.log(
                recipient.addrType, // "EX", "SMTP"
                recipient.emailAddress,
                recipient.displayName
            );
        }

        const attachments = await message.getAttachments();
        for (const att of attachments) {
            console.log(
                att.attachMethod,
                att.displayName,
                att.filename
            );

            if (att.attachMethod === PSTAttachment.ATTACHMENT_METHOD_BY_VALUE) {
                console.log(att.fileData);
            }

            if (att.attachMethod === PSTAttachment.ATTACHMENT_METHOD_EMBEDDED) {
                const embeddedMessage = await att.getEmbeddedPSTMessage();
                console.log(embeddedMessage);
            }
        }
    }

    const subFolders = await folder.getSubFolders();
    for (const subFolder of subFolders) {
        await traverseFolder(subFolder);
    }
}

const rootFolder = await pstFile.getRootFolder();
traverseFolder(rootFolder);
```

## Authors

Original repository's author:

- Ed Pfromer (epfromer@gmail.com)

Forked repository's author:

- kenjiuno

## License

MIT © [epfromer](https://github.com/epfromer), kenjiuno
