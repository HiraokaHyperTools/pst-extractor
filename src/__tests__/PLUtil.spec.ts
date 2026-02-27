import { openLowPst } from '../PLUtil.js';
import { getPropertyContext } from '../PropertyContextUtil.js';
import { PropertyValueResolverV1 } from '../PropertyValueResolverV1.js';
import { processNameToIDMap } from '../PAUtil.js';
import { openPstFile, PSTFolder } from '../index.js';
import fs from 'fs';
import { getHeapFrom } from '../PHUtil';
import { resolve } from 'path';
import { PSTUtil } from '../PSTUtil.class.js';

// cls & yarn test:unit --testMatch "**/PLUtil.spec.ts" --coverage false

const filePath = resolve('./src/__tests__/testdata/mtnman1965@outlook.com.ost');

describe('PLUtil/PHUtil tests', () => {
  async function doOpenPstFile(filePath: string) {
    const pst = await openPstFile(filePath);

    async function walk(folder: PSTFolder, depth: number): Promise<void> {
      for (let item of await folder.getEmails()) {
        //console.log("  ".repeat(depth) + `- ${item.displayName}`);
      }
      for (let subFolder of await folder.getSubFolders()) {
        //console.log("  ".repeat(depth) + `- ${subFolder.displayName}`);
        await walk(subFolder, depth + 1);
      }
    }

    await walk(await pst.getRootFolder(), 0);

    await pst.close();
  }

  it('openPstFile ost', async () => await doOpenPstFile(filePath));

  it('openPstFile pst unicode', async () => await doOpenPstFile(
    resolve('./src/__tests__/testdata/contacts.pst'))
  );

  it('openPstFile pst ansi', async () => await doOpenPstFile(
    resolve('./src/__tests__/testdata/contacts97-2002.pst'))
  );

  it('openLowPst and so on', async () => {
    const file = await fs.promises.open(filePath, "r");
    try {
      async function readFile(
        buffer: ArrayBuffer, offset: number, length: number, position: number
      ): Promise<number> {
        const view = new Uint8Array(buffer);
        const { bytesRead } = await file.read(view, offset, length, position);
        return bytesRead;
      }

      const lowPst = await openLowPst({ readFile, close: async () => { } });

      const resolver = new PropertyValueResolverV1(
        PSTUtil.createConvertAnsiString("ms932")
      );

      {
        const nodeMap = await processNameToIDMap(
          await lowPst.getOneNodeByOrError(97),
          resolver
        );

        //console.log(nodeMap);
      }

      const rootNode = await lowPst.getOneNodeByOrError(290);

      const heap = await getHeapFrom(rootNode.getSubNode());
      const pcBufs = await heap.getReader().getHeapBuffers(heap.userRootHnid);
      //console.log({ pcBufs });

      const pc = await getPropertyContext(
        heap,
        resolver
      );

      {
        const nameToIdMapDescriptorNode = await lowPst.getOneNodeByOrError(97);
        //console.log({ nameToIdMapDescriptorNode });
        const nameToIdMapHeap = await getHeapFrom(nameToIdMapDescriptorNode.getSubNode());
        //console.log({ nameToIdMap });
        const more = await getPropertyContext(
          nameToIdMapHeap,
          resolver
        );
        //console.log(await more.list());
      }

      // {
      //   const alphaNode = await lowPst.getOneNodeByOrError(2097188);
      //   const alphaHeap = await getHeapFromSub(alphaNode.getNodeReader(), 0x671);
      //   const alphaMore = await getTableContext(
      //     alphaHeap,
      //     resolver
      //   );
      //   //console.log(await (await alphaMore.rows())[0].list());
      // }

      //console.log(await pc.list());
    } finally {
      await file.close();
    }
  })
})
