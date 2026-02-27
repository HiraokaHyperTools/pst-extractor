import { copyFile } from "fs";
import { getBTHeapReaderFrom } from "./BTHeap.js";
import type { PHNodeHeap } from "./PHNodeHeap.js";
import { splitPer } from "./PLMisc.js";
import type { Property } from "./Property.js";
import type { PropertyValueResolver } from "./PropertyValueResolver.js";
import type { RawProperty } from "./RawProperty.js";
import type { TableContext } from "./TableContext.js";
import type { TableRow } from "./TableRow.js";

const bTypeTC = 0x7c;

function copy<T>(rows: T[]): T[] {
  return rows.map(
    it => Object.assign({}, it)
  )
}

interface Record {
  buffer: ArrayBuffer,
  ceb: boolean[],
}

/**
 * TCOLDESC
 * 
 * @see [[MS-PST]: TCOLDESC | Microsoft Docs](https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-pst/3a2f63cf-bb40-4559-910c-e55ec43d9cbb)
 */
interface Column {
  type: number;
  key: number;
  ibData: number;
  cbData: number;
  iBit: number;
}

export async function getTableContext(
  heap: PHNodeHeap,
  resolver: PropertyValueResolver
): Promise<TableContext> {
  if (heap.bClientSig != bTypeTC) {
    throw new Error("expected type bTypeTC");
  }

  const reader = heap.getReader();

  const headerData = await reader.getHeapBuffers(heap.userRootHnid);
  if (headerData.length !== 1) {
    throw new Error("table context header buffer must be single");
  }
  const headerView = new DataView(headerData[0]);

  const tcSig = headerView.getUint8(0);
  const numCols = headerView.getUint8(1);
  const cebOffset = headerView.getUint16(6, true);
  const numRowBytes = headerView.getUint16(8, true);
  const hidRowIndex = headerView.getUint32(10, true);
  const hnidRows = headerView.getUint32(14, true);

  var tColDesc = headerData[0].slice(22);

  const schema: Column[] = Array.from(splitPer(tColDesc, 8))
    .map(
      it => {
        const view = new DataView(it);
        return {
          type: view.getUint16(0, true),
          key: view.getUint16(2, true),
          ibData: view.getUint16(4, true),
          cbData: view.getUint8(6),
          iBit: view.getUint8(7),
        } as Column;
      }
    );

  if (numCols !== schema.length) {
    throw new Error("schema length not matched");
  }
  if (tcSig !== 0x7c) {
    throw new Error("tcSig is not 0x7c");
  }

  //const min_size = schema.reduce((accum, it) => accum + it.cbData, 0);

  //const rowIndex = getBTHeapReaderFrom(reader, hidRowIndex);

  const rows_pages = (hnidRows !== 0)
    ? await reader.getHeapBuffers(hnidRows)
    : [];

  //const data2 = await reader.getHeapBuffers(offset2);

  const rows_per_page = (rows_pages.length !== 0)
    ? Math.floor(rows_pages[0].byteLength / numRowBytes)
    : 0;

  function get_record(record_index: number): Record {
    const page_index = (record_index / rows_per_page) | 0;
    const heap_index = (record_index % rows_per_page) | 0;

    const buffer = rows_pages[page_index].slice(
      numRowBytes * (heap_index + 0),
      numRowBytes * (heap_index + 1)
    );
    if (buffer.byteLength !== numRowBytes) {
      throw new Error(`get_record(${record_index}) is reaching EOF while reading record. The heap is ${reader}`);
    }

    const ceb: boolean[] = [];
    {
      const rgCEB = new Uint8Array(buffer.slice(cebOffset));
      for (let x = 0; x < numCols; x++) {
        ceb.push((rgCEB[(x / 8) | 0] & (1 << (7 - (x % 8)))) != 0);
      }
    }

    return {
      buffer,
      ceb,
    };
  }

  const count = rows_pages.reduce((accum, it) => (accum + it.byteLength / numRowBytes) | 0, 0);

  async function listRaw(index: number): Promise<RawProperty[]> {
    const record = get_record(index);
    const list: RawProperty[] = [];
    for (let column of schema) {
      list.push({
        key: column.key,
        type: column.type,
        value: record.ceb[column.iBit]
          ? record.buffer.slice(column.ibData, column.ibData + column.cbData)
          : new ArrayBuffer(0),
      });
    }
    return list;
  }

  async function list(record: number): Promise<Property[]> {
    const rawProps = await listRaw(record);
    const list: Property[] = [];
    for (let rawProp of rawProps) {
      try {
        list.push({
          key: rawProp.key,
          type: rawProp.type,
          value: (rawProp.value.byteLength !== 0)
            ? await resolver.resolveValueOf(
              rawProp.key,
              rawProp.type,
              rawProp.value,
              reader
            )
            : undefined
        });
      }
      catch (ex) {
        throw new Error(`getTableContext.list(rowIndex=${record}) resolving property`
          + ` key=0x${rawProp.key.toString(16).padStart(4, '0')}`
          + ` type=0x${rawProp.type.toString(16).padStart(4, '0')}`
          + ` of ${heap} failure`
          + ` --> ${ex}`
        );
      }
    }
    return list;
  }

  const rows: TableRow[] = [];
  for (let x = 0; x < count; x++) {
    rows.push({
      async listRaw(): Promise<RawProperty[]> {
        return await listRaw(x);
      },
      async list(): Promise<Property[]> {
        return await list(x);
      },
    });
  }

  return {
    async rows(): Promise<TableRow[]> {
      return copy(rows);
    },
  } as TableContext;
}
