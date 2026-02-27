import type { PHNodeHeapReader } from "./PHNodeHeapReader.js";

export async function getBTHeapReaderFrom(
  heap: PHNodeHeapReader,
  hnid: number
): Promise<BTHeapReader> {
  const header = await heap.getHeapBuffers(hnid);
  if (header.length !== 1) {
    throw new Error(`btree heap buffer must be single`);
  }
  const headerView = new DataView(header[0]);
  const bTypeBTH = 0xB5;
  if (headerView.getUint8(0) !== bTypeBTH) {
    throw new Error(`signature must be bTypeBTH`);
  }
  const cbKey = headerView.getUint8(1);
  const cbEnt = headerView.getUint8(2);
  const bIdxLevels = headerView.getUint8(3);
  const hidRoot = headerView.getUint32(4, true);

  const list: BTHeapEntry[] = [];

  async function recursive(hid: number, level: number) {
    if (hid !== 0) {
      if (level === 0) {
        const recordSize = cbKey + cbEnt;
        const records = (await heap.getHeapBuffers(hid));
        if (records.length !== 1) {
          throw new Error(`btree heap record must be single`);
        }
        const record = records[0];
        const numRecords = Math.floor(record.byteLength / recordSize);
        for (let x = 0; x < numRecords; x++) {
          const top = recordSize * x;
          list.push({
            key: record.slice(top, top + cbKey),
            data: record.slice(top + cbKey, top + recordSize),
          });
        }
      }
      else {
        const recordSize = cbKey + 4;
        const records = (await heap.getHeapBuffers(hid));
        if (records.length !== 1) {
          throw new Error(`btree intermediate record must be single`);
        }
        const record = records[0];
        const recordView = new DataView(record);
        const numRecords = Math.floor(record.byteLength / recordSize);
        for (let x = 0; x < numRecords; x++) {
          const top = recordSize * x;
          const hidInner = recordView.getUint32(top + cbKey, true);
          await recursive(hidInner, level - 1);
        }
      }
    }
  }

  await recursive(hidRoot, bIdxLevels);

  return {
    async list() {
      return list;
    },
  } as BTHeapReader;
}

export interface BTHeapReader {
  list(): Promise<BTHeapEntry[]>;
}

export interface BTHeapEntry {
  /**
   * 2, 4, 8, or 16
   */
  key: ArrayBuffer;

  /**
   * 1 ~ 32
   */
  data: ArrayBuffer;
}
