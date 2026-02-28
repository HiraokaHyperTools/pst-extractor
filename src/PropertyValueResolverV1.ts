import Long from "long";
import type { PHNodeHeapReader } from "./PHNodeHeapReader.js";
import { splitPer } from "./PLMisc.js";
import { readLong } from "./PLUtil.js";
import { PropertyTypeObject } from "./PropertyTypeObject.js";
import type { PropertyValueResolver } from "./PropertyValueResolver.js";
import { PSTUtil } from "./PSTUtil.class.js";

export interface PrimitiveTypeConverterArg {
  /**
   * Access to raw property value.
   * 
   * Do not modify the contents.
   */
  view: DataView;

  /**
   * Access to heap reader.
   */
  heap: PHNodeHeapReader;

  /**
   * Get bytes from either raw property value or heap.
   */
  getBytes: (numBytes: number) => Promise<ArrayBuffer | undefined>;

  /**
   * Resolve heap by hnid.
   */
  resolveHeap: (heap: number) => Promise<ArrayBuffer | undefined>;

  /**
   * Convert ansiString to unicode string.
   */
  convertAnsiStringImmediately: (data: Uint8Array) => string;
}

/**
 * @example
 * 
 * ```ts
 * async function convertShortProperty(arg: PrimitiveTypeConverterArg): Promise<any> {
 *  return arg.view.getInt16(0, true);
 * }
 * ```
 * 
 * ```ts
 * async function convertAnsiStringProperty(arg: PrimitiveTypeConverterArg): Promise<any> {
 *   const heap = arg.view.getUint32(0, true);
 *   const bytes = await arg.resolveHeap(heap);
 *   return (bytes !== undefined)
 *     ? arg.convertAnsiString(new Uint8Array(bytes))
 *     : undefined;
 * }
 * ```
 */
export type PrimitiveTypeConverter = (
  arg: PrimitiveTypeConverterArg
) => Promise<any>;

type PrimitiveTypeConverters = { [key: number]: PrimitiveTypeConverter };

const typeConverters: PrimitiveTypeConverters = {};

const PT_BOOLEAN = 0xB;
const PT_DOUBLE = 0x5;
const PT_LONG = 0x3;
const PT_OBJECT = 0xD;
const PT_LONGLONG = 0x14;
const PT_STRING8 = 0x1E;
const PT_UNICODE = 0x1F;
const PT_SYSTIME = 0x40;
const PT_CLSID = 0x48;
const PT_SHORT = 0x2;
const PT_FLOAT = 0x4;

const PT_MV_UNICODE = 0x101F;
const PT_MV_STRING8 = 0x101E;
const PT_MV_BINARY = 0x1102;
const PT_MV_LONG = 0x1003;
const PT_MV_CLSID = 0x1048;
const PT_MV_SHORT = 0x1002;

const PT_MVPV_BINARY = 0x2102;

typeConverters[PT_SHORT] = async (arg) => {
  return arg.view.getInt16(0, true);
};
typeConverters[PT_FLOAT] = async (arg) => {
  return arg.view.getFloat32(0, true);
};
typeConverters[PT_LONG] = async (arg) => {
  return arg.view.getInt32(0, true);
};
typeConverters[PT_OBJECT] = async (arg) => {
  const heap = arg.view.getUint32(0, true);
  const bytes = await arg.resolveHeap(heap);
  if (bytes !== undefined) {
    const view = new DataView(bytes);
    return new PropertyTypeObject(
      view.getUint32(0, true),
      view.getUint32(4, true)
    );
  }
  return bytes;
};
typeConverters[PT_LONGLONG] = async (arg) => {
  const bytes = await arg.getBytes(8);
  if (bytes !== undefined) {
    const view = new DataView(bytes);
    return readLong(view, 0);
  }
  else {
    return undefined;
  }
};
typeConverters[PT_DOUBLE] = async (arg) => {
  const bytes = await arg.getBytes(8);
  return (bytes !== undefined)
    ? new DataView(bytes).getFloat64(0, true)
    : undefined;
};
typeConverters[PT_BOOLEAN] = async (arg) => {
  return arg.view.getUint8(0) !== 0;
};
typeConverters[PT_STRING8] = async (arg) => {
  const heap = arg.view.getUint32(0, true);
  const bytes = await arg.resolveHeap(heap);
  return (bytes !== undefined)
    ? arg.convertAnsiStringImmediately(new Uint8Array(bytes))
    : undefined;
};
typeConverters[PT_UNICODE] = async (arg) => {
  const heap = arg.view.getUint32(0, true);
  const bytes = await arg.resolveHeap(heap);
  return (bytes !== undefined)
    ? PSTUtil.decodeUtf16leString(new Uint8Array(bytes)).replace(/\0/g, '')
    : undefined;
  // `.replace(/\0/g, '')` is needed to eliminate a trailing null char.
};
typeConverters[PT_SYSTIME] = async (arg) => {
  const bytes = await arg.getBytes(8);
  if (bytes !== undefined) {
    const view = new DataView(bytes);
    return PSTUtil.filetimeToDate(
      new Long(view.getUint32(4, true)),
      new Long(view.getUint32(0, true))
    );
  }
  else {
    return undefined;
  }
};
typeConverters[PT_CLSID] = async (arg) => {
  const heap = arg.view.getUint32(0, true);
  const bytes = await arg.resolveHeap(heap);
  return bytes;
};
typeConverters[0x0102] = async (arg) => {
  const heap = arg.view.getUint32(0, true);
  const bytes = await arg.resolveHeap(heap);
  return bytes;
};
typeConverters[PT_MV_UNICODE] = async (arg) => {
  const heap = arg.view.getUint32(0, true);
  const list = [] as any[];
  if (heap !== 0) {
    const bytes = await arg.resolveHeap(heap);
    if (bytes !== undefined) {
      const view = new DataView(bytes);
      const count = view.getUint32(0, true);
      for (let x = 0; x < count - 1; x++) {
        const from = view.getUint32(4 + 4 * (x), true);
        const to = view.getUint32(4 + 4 * (x + 1), true);

        const elementBytes = bytes.slice(from, to);

        list.push(
          PSTUtil.decodeUtf16leString(new Uint8Array(elementBytes)).replace(/\0/g, '')
        )
      }
    }
  }
  return list;
};
typeConverters[PT_MV_BINARY] = async (arg) => {
  const heap = arg.view.getUint32(0, true);
  const list = [] as any[];
  if (heap !== 0) {
    const bytes = await arg.resolveHeap(heap);
    if (bytes !== undefined) {
      const view = new DataView(bytes);
      const count = view.getUint32(0, true);
      for (let x = 0; x < count - 1; x++) {
        const from = view.getUint32(4 + 4 * (x), true);
        const to = view.getUint32(4 + 4 * (x + 1), true);

        const elementBytes = bytes.slice(from, to);

        list.push(elementBytes)
      }
    }
  }
  return list;
};
typeConverters[PT_MV_LONG] = async (arg) => {
  const heap = arg.view.getUint32(0, true);
  const list = [] as any[];
  if (heap !== 0) {
    const bytes = await arg.resolveHeap(heap);
    if (bytes !== undefined) {
      const view = new DataView(bytes);
      const count = bytes.byteLength / 4;
      for (let x = 0; x < count; x++) {
        list.push(view.getInt32(4 * x, true))
      }
    }
  }
  return list;
};
typeConverters[PT_MV_CLSID] = async (arg) => {
  const heap = arg.view.getUint32(0, true);
  const list = [] as any[];
  if (heap !== 0) {
    const bytes = await arg.resolveHeap(heap);
    if (bytes !== undefined) {
      const count = bytes.byteLength / 16;
      for (let x = 0; x < count; x++) {
        const from = 16 * (x);
        const to = 16 * (x + 1);

        const elementBytes = bytes.slice(from, to);

        list.push(elementBytes)
      }
    }
  }
  return list;
};
typeConverters[PT_MV_SHORT] = async (arg) => {
  const heap = arg.view.getUint32(0, true);
  const list = [] as any[];
  if (heap !== 0) {
    const bytes = await arg.resolveHeap(heap);
    if (bytes !== undefined) {
      const view = new DataView(bytes);
      const count = bytes.byteLength / 2;
      for (let x = 0; x < count; x++) {
        list.push(view.getInt16(2 * x, true))
      }
    }
  }
  return list;
};
typeConverters[PT_MV_STRING8] = async (arg) => {
  const heap = arg.view.getUint32(0, true);
  const list = [] as any[];
  if (heap !== 0) {
    const bytes = await arg.resolveHeap(heap);
    if (bytes !== undefined) {
      const view = new DataView(bytes);
      const count = view.getUint32(0, true);
      for (let x = 0; x < count - 1; x++) {
        const from = view.getUint32(4 + 4 * (x), true);
        const to = view.getUint32(4 + 4 * (x + 1), true);

        list.push(
          await arg.convertAnsiStringImmediately(new Uint8Array(bytes, from, to - from))
        )
      }
    }
  }
  return list;
};
typeConverters[PT_MVPV_BINARY] = async (arg) => {
  // not sure
  const heap = arg.view.getUint32(0, true);
  const list = [] as any[];
  if (heap !== 0) {
    const bytes = await arg.resolveHeap(heap);
    if (bytes !== undefined) {
      const view = new DataView(bytes);
      const count = view.getUint32(0, true);
      for (let x = 0; x < count - 1; x++) {
        const from = view.getUint32(4 + 4 * (x), true);
        const to = view.getUint32(4 + 4 * (x + 1), true);

        const elementBytes = bytes.slice(from, to);

        list.push(elementBytes)
      }
    }
  }
  return list;
};


function mixIntoOne(array: ArrayBuffer[]): ArrayBuffer {
  if (array.length === 0) {
    return new ArrayBuffer(0);
  }
  else if (array.length === 1) {
    return array[0];
  }
  else {
    const numBytes = array.reduce((prev, it) => prev + it.byteLength, 0);
    const one = new ArrayBuffer(numBytes);
    const dest = new Uint8Array(one);
    array.reduce(
      (nextPos, source) => {
        dest.set(new Uint8Array(source), nextPos);
        return nextPos + source.byteLength;
      },
      0
    );
    return one;
  }
}

/**
 * @internal
 */
export class PropertyValueResolverV1 implements PropertyValueResolver {
  private convertAnsiString: (data: Uint8Array) => string;
  private provideTypeConverterOf: ((type: number) => PrimitiveTypeConverter | undefined);

  constructor(
    convertAnsiString: (data: Uint8Array) => string,
    provideTypeConverterOf?: (type: number) => PrimitiveTypeConverter | undefined,
    provideFallbackTypeConverterOf?: (type: number) => PrimitiveTypeConverter | undefined
  ) {
    this.convertAnsiString = convertAnsiString;
    this.provideTypeConverterOf = type => {
      if (provideTypeConverterOf) {
        const converter = provideTypeConverterOf(type);
        if (converter) {
          return converter;
        }
      }
      const converter = typeConverters[type];
      if (converter) {
        return converter;
      }
      if (provideFallbackTypeConverterOf) {
        const converter = provideFallbackTypeConverterOf(type);
        if (converter) {
          return converter;
        }
      }
      return undefined;
    };
  }

  async resolveValueOf(
    key: number,
    type: number,
    value: ArrayBuffer,
    heap: PHNodeHeapReader
  ): Promise<any> {
    const view = new DataView(value);

    async function resolveHeap(hnid: number): Promise<ArrayBuffer | undefined> {
      return mixIntoOne(
        await heap.getHeapBuffers(hnid)
      );
    }

    const converter = this.provideTypeConverterOf(type);
    if (converter === undefined || converter === null) {
      throw new Error(`property type 0x${type.toString(16)} is unknown. please define a typeConverter in PropertyValueResolverV1.ts`);
    }

    return await converter(
      {
        view: view,
        heap: heap,
        async getBytes(numBytes: number): Promise<ArrayBuffer | undefined> {
          const { byteLength } = value;
          if (byteLength === 4 && byteLength < numBytes) {
            const hnid = view.getUint32(0, true);
            return await resolveHeap(hnid);
          }
          else if (byteLength !== 4 && numBytes <= byteLength) {
            return value.slice(0, numBytes);
          }
          else {
            throw new Error("dont know how to provide buffer");
          }
        },
        resolveHeap: resolveHeap,
        convertAnsiStringImmediately: this.convertAnsiString,
      }
    );
  }
}
