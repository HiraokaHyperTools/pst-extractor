import { PSTFile } from './PSTFile.class.js';
import fs from 'fs';
import { openLowPst, type ReadFileApi } from './PLUtil.js';
import { PropertyValueResolverV1 } from './PropertyValueResolverV1.js';
import { processNameToIDMap } from './PAUtil.js';
import type { PSTOpts } from './PSTOpts.js';
import { PSTUtil } from './PSTUtil.class.js';
import type { IPSTFile } from './IPSTFile.js';

/**
 * Open pst/ost file from os file path.
 * 
 * @param path os file path.
 * @param opts options
 * @returns 
 */
export async function openPstFile(
  path: string,
  opts?: PSTOpts
): Promise<IPSTFile> {
  const file = await fs.promises.open(path, "r");

  async function readFile(
    buffer: ArrayBuffer, offset: number, length: number, position: number
  ): Promise<number> {
    const view = new Uint8Array(buffer);
    const { bytesRead } = await file.read(view, offset, length, position);
    return bytesRead;
  }

  return await openPst(
    {
      readFile,
      close: async (): Promise<void> => {
        await file.close();
      },
    },
    opts
  );
}

/**
 * Open pst/ost file using user defined callback.
 * 
 * @param api reader callback
 * @param opts options
 * @returns 
 */
export async function openPst(
  api: ReadFileApi,
  opts?: PSTOpts
): Promise<IPSTFile> {
  const lowPst = await openLowPst(api);

  const convertAnsiString = (opts && opts.convertAnsiString)
    || PSTUtil.createConvertAnsiString(
      (opts && opts.ansiEncoding) || "latin1"
    );

  const resolver = new PropertyValueResolverV1(
    convertAnsiString,
    opts?.provideTypeConverterOf,
    opts?.provideFallbackTypeConverterOf
  );

  const nodeMap = await processNameToIDMap(
    await lowPst.getOneNodeByOrError(97),
    resolver
  );

  return new PSTFile(lowPst, nodeMap, opts);
}
