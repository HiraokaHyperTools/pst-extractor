import { PSTFile } from '../PSTFile.class.js';
import { PSTFolder } from '../PSTFolder.class.js';
import { openPstFile } from '../openPstFile.js';
import { resolve, extname } from 'path';
import fs from 'fs';

// yarn test:unit --coverage false -t "tree traversal tests"

// add dir list having pst/ost files
const sourceDirs: string[] = [
  // './src/__tests__/testdata/',
];

// add file list of pst/ost files
const sourceFiles: string[] = [
  // './src/__tests__/testdata/mtnman1965@outlook.com.ost',
];

describe('tree traversal tests', () => {
  sourceDirs.forEach(
    dir => {
      for (let name of fs.readdirSync(resolve(dir))) {
        const fileExt = extname(name).toLowerCase();
        if (fileExt === ".pst" || fileExt === ".ost") {
          sourceFiles.push(resolve(dir, name));
        }
      }
    }
  );

  if (sourceFiles.length !== 0) {
    it.each(sourceFiles)("scan %s", async (sourceFile: string) => {
      await scan(sourceFile);
    }, 1000 * 60 * 60);

    it.each(sourceFiles)("faster scan %s", async (sourceFile: string) => {
      await fasterScan(sourceFile);
    }, 1000 * 60 * 60);
  }
  else {
    it("no files given", function () { });
  }
});

async function scan(sourceFile: string): Promise<void> {
  const pstFile = await openPstFile(sourceFile);
  try {
    await scanTree((await pstFile.getRootFolder()), 0);
  }
  finally {
    pstFile.close();
  }
}

async function scanTree(folder: PSTFolder, depth: number): Promise<void> {
  function prefix(text: string): string {
    return "  ".repeat(depth) + text;
  }
  for (let subFolder of (await folder.getSubFolders())) {
    //console.log(prefix(`@ ${subFolder.displayName}`));
    await scanTree(subFolder, depth + 1);
  }
  for (let item of (await folder.getEmails())) {
    //console.log(prefix(`- ${item.displayName}`));
    (await item.getAttachments());
    (await item.getRecipients());
  }
}

async function fasterScan(sourceFile: string): Promise<void> {
  const pstFile = await openPstFile(sourceFile);
  try {
    await fasterScanTree((await pstFile.getRootFolder()), 0);
  }
  finally {
    pstFile.close();
  }
}

async function fasterScanTree(folder: PSTFolder, depth: number): Promise<void> {
  function prefix(text: string): string {
    return "  ".repeat(depth) + text;
  }
  for (let subFolder of (await folder.getSubFolders())) {
    //console.log(prefix(`@ ${subFolder.displayName}`));
    await scanTree(subFolder, depth + 1);
  }
  for (let faster of (await folder.getFasterEmailList())) {
    const item = await faster.getMessage();
    //console.log(prefix(`- ${item.displayName}`));
    (await item.getAttachments());
    (await item.getRecipients());
  }
}
