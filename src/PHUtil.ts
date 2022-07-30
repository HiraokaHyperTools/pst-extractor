/**
 * PST higher level utilities
 */

import { PHNodeHeap } from "./PHNodeHeap";
import { PHNodeHeapReader } from "./PHNodeHeapReader";
import { PLNodeReader } from "./PLNodeReader";

const USE_MAIN_DATA = -1;

export async function getHeapFromMain(node: PLNodeReader): Promise<PHNodeHeap> {
  return getHeap(node, USE_MAIN_DATA);
}

export async function getHeapFromSub(node: PLNodeReader, subNodeId: number): Promise<PHNodeHeap> {
  return getHeap(node, subNodeId);
}

async function getHeap(node: PLNodeReader, subNodeId: number): Promise<PHNodeHeap> {
  const data_array = (subNodeId === USE_MAIN_DATA)
    ? [await node.getMainData()]
    : await node.getSubDataArray(subNodeId);
  const data_chunks = new Map<number, ArrayBuffer>();
  let bClientSig = 0;
  let userRootHnid = 0;

  for (let x = 0; x < data_array.length; x++) {
    if (x === 0) {
      ({ bClientSig, userRoot: userRootHnid } = load_root_header(data_chunks, data_array[x]));
    }
    else {
      load_page_header(data_chunks, data_array[x], x);
    }
  }

  return {
    bClientSig,
    userRootHnid,

    getReader() {
      return {
        async getHeapBuffers(hnid): Promise<ArrayBuffer[]> {
          if (hnid === 0) {
            return [];
          }
          else if (hnid & 0x1f) {
            // this is NID (node)
            const data_array = await node.getSubDataArray(hnid);
            return data_array;
          }
          else {
            // this is HID (heap)
            const data = data_chunks.get(hnid);
            if (data === undefined) {
              throw new Error("heap not found");
            }
            return [data];
          }
        },
      } as PHNodeHeapReader;
    },
  } as PHNodeHeap;
}

/**
 * Parse HNHDR
 * 
 * @see https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-pst/8e4ae05c-3c24-4103-b7e5-ffef6f244834
 */
function load_root_header(
  data_chunks: Map<number, ArrayBuffer>,
  data: ArrayBuffer
): { bClientSig: number, userRoot: number } {
  const view = new DataView(data);
  const page_map = view.getUint16(0, true);
  const sig = view.getUint8(2);
  const bClientSig = view.getUint8(3);
  const userRoot = view.getUint32(4, true);

  if (sig !== 0xec) {
    throw new Error("invalid HNHDR signature found");
  }

  // read HNPAGEMAP
  const offsets_count = view.getUint16(page_map, true) + 1;
  for (let x = 0; x < offsets_count - 1; x++) {
    const from = view.getUint16(page_map + 4 + 2 * (x), true);
    const to = view.getUint16(page_map + 4 + 2 * (x + 1), true);

    data_chunks.set(0x20 * (1 + x), data.slice(from, to));
  }

  return { bClientSig, userRoot };
}

function load_page_header(
  data_chunks: Map<number, ArrayBuffer>,
  data: ArrayBuffer,
  page_index: number
): void {
  const view = new DataView(data);
  const page_map = view.getUint16(0, true);

  const offsets_count = view.getUint16(page_map, true) + 1;
  for (let x = 0; x < offsets_count - 1; x++) {
    const from = view.getUint16(page_map + 4 + 2 * (x), true);
    const to = view.getUint16(page_map + 4 + 2 * (x + 1), true);

    data_chunks.set(0x20 * (1 + x) + 65536 * page_index, data.slice(from, to));
  }
}
