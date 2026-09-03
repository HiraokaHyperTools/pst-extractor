/**
 * PST lower level utilities
 */

import Long from "long";
import type { PLNode } from "./PLNode.js";
import type { PLStore } from "./PLStore.js";
import type { PLSubNode } from "./PLSubNode.js";
import { PSTUtil } from "./PSTUtil.class.js";
import { inflate } from "pako";

export type ReadFile = (
  buffer: ArrayBuffer,
  offset: number,
  length: number,
  position: number
) => Promise<number>;

export type Close = () => Promise<void>;

/**
 * Defines a callback based async file operation API.
 */
export interface ReadFileApi {
  /**
   * Defines read API like `fs.promises.readFile`.
   */
  readFile: ReadFile,

  /**
   * Defines close API like `fs.promises.close`.
   */
  close: Close,
}

function surelyReader(readFile: ReadFile): ReadFile {
  return async (buffer: ArrayBuffer, offset: number, length: number, position: number): Promise<number> => {
    const bytesRead = await readFile(buffer, offset, length, position);
    if (bytesRead !== length) {
      throw new Error(`file reader must provide buffer ${length} bytes from position ${position}, while we got ${bytesRead} bytes`);
    }
    return length;
  };
}

interface NodeFooter {
  itemCount: number;
  level: number;
  thisId: Long;
}

interface XBlockPtr {
  blockId: number;
}

interface XXBlockPtr {
  blockId: number;
}

interface StoreTrait {
  readXBlock(view: DataView, offset: number, itemCount: number): XBlockPtr[];
  readXXBlock(view: DataView, offset: number, itemCount: number): XBlockPtr[];
  readSIBlock(view: DataView, offset: number): SIBlock;
  readSLBlock(view: DataView, offset: number): SLBlock;
  readNodePtr(view: DataView, offset: number, footer: NodeFooter): NodePtr[];
  readTablePtr(view: DataView, offset: number, footer: NodeFooter): TablePtr[];
  readBlockPtr(view: DataView, offset: number, footer: NodeFooter): BlockPtr[];
  BACKLINK_OFFSET: number;
  LEVEL_INDICATOR_OFFSET: number;
  ITEM_COUNT_OFFSET: number;
  readId(view: DataView, offset: number): Long;

  ENC_OFFSET: number;
  SECOND_POINTER_COUNT: number;
  SECOND_POINTER: number;
  INDEX_POINTER_COUNT: number;
  INDEX_POINTER: number;
  BlockSize: number;
  unzipHook: UnzipHook;
  is4K: boolean;
  /**
   * Used when exporting a message to msg file.
   * 
   * - Setting this value to `PidTagStoreSupportMask`.
   * - Set `undefined` if `PidTagStoreSupportMask` should not be emitted.
   */
  storeSupportMask: number | undefined;
}

/**
 * @internal
 */
export function readNumber64(view: DataView, offset: number): number {
  return new Long(
    view.getUint32(offset, true),
    view.getUint32(offset + 4, true),
    true
  )
    .toNumber();
}

/**
 * @internal
 */
export function readLong(view: DataView, offset: number): Long {
  return new Long(
    view.getUint32(offset, true),
    view.getUint32(offset + 4, true),
    true
  );
}

type UnzipHook = (data: ArrayBuffer) => Promise<ArrayBuffer>;

/**
 * 
 * @internal
 */
async function passThru1(data: ArrayBuffer): Promise<ArrayBuffer> {
  return data;
}

/**
 * 
 * @internal
 */
async function willUnzip1(data: ArrayBuffer): Promise<ArrayBuffer> {
  if (data.byteLength >= 4) {
    const view = new DataView(data);
    if (view.getUint16(0, true) === 0x9c78) {
      const inflated = inflate(data);
      return inflated.buffer as ArrayBuffer;
    }
  }

  return data;
}


namespace ptr32 {
  export function readBlockPtr(
    view: DataView,
    offset: number,
    footer: NodeFooter
  ): BlockPtr[] {
    const array: BlockPtr[] = [];
    for (let x = 0; x < footer.itemCount; x++) {
      const top = offset + 12 * x;
      const blockId = view.getUint32(top + 0, true);
      array.push(
        {
          blockId,
          offset: new Long(view.getUint32(top + 4, true)),
          size: view.getUint16(top + 8, true),
          isData: (blockId & 2) === 0,
        } as BlockPtr
      );
    }
    return array;
  }

  export function readTablePtr(
    view: DataView,
    offset: number,
    footer: NodeFooter
  ): TablePtr[] {
    const array: TablePtr[] = [];
    for (let x = 0; x < footer.itemCount; x++) {
      const top = offset + 12 * x;
      array.push(
        {
          start: new Long(view.getUint32(top + 0, true)),
          u1: new Long(view.getUint32(top + 4, true)),
          offset: new Long(view.getUint32(top + 8, true)),
        } as TablePtr
      );
    }
    return array;
  }

  export function readNodePtr(
    view: DataView,
    offset: number,
    footer: NodeFooter
  ): NodePtr[] {
    const array: NodePtr[] = [];
    for (let x = 0; x < footer.itemCount; x++) {
      const top = offset + 16 * x;
      array.push(
        {
          nodeId: view.getUint32(top + 0, true),
          blockId: view.getUint32(top + 4, true),
          subBlockId: view.getUint32(top + 8, true),
          parentNodeId: view.getUint32(top + 12, true),
        } as NodePtr
      );
    }
    return array;
  }

  export function readSLBlock(
    view: DataView,
    offset: number
  ): SLBlock {
    const count = view.getUint16(offset + 2, true);
    const entries: SLBlockEntry[] = [];
    for (let x = 0; x < count; x++) {
      const top = offset + 4 + 12 * x;
      entries.push(
        {
          nodeId: view.getUint32(top + 0, true),
          blockId: view.getUint32(top + 4, true),
          subBlockId: view.getUint32(top + 8, true),
        } as SLBlockEntry
      );
    }
    return {
      entries
    };
  }

  export function readXBlock(
    view: DataView,
    offset: number,
    itemCount: number
  ): XBlockPtr[] {
    const entries: XBlockPtr[] = [];
    for (let x = 0; x < itemCount; x++) {
      const top = offset + 8 + 4 * x;
      entries.push(
        {
          blockId: view.getUint32(top + 0, true),
        } as XBlockPtr
      );
    }
    return entries;
  }

  export function readXXBlock(
    view: DataView,
    offset: number,
    itemCount: number
  ): XXBlockPtr[] {
    const entries: XXBlockPtr[] = [];
    for (let x = 0; x < itemCount; x++) {
      const top = offset + 8 + 4 * x;
      entries.push(
        {
          blockId: view.getUint32(top + 0, true),
        } as XXBlockPtr
      );
    }
    return entries;
  }

  export function readSIBlock(
    view: DataView,
    offset: number
  ): SIBlock {
    const count = view.getUint16(offset + 2, true);
    const entries: SIBlockEntry[] = [];
    for (let x = 0; x < count; x++) {
      const top = offset + 4 + 12 * x;
      entries.push(
        {
          nodeId: view.getUint32(top + 0, true),
          blockId: view.getUint32(top + 4, true),
        } as SIBlockEntry
      );
    }
    return {
      entries
    };
  }
}


namespace ptr64 {
  export function readBlockPtr(
    view: DataView,
    offset: number,
    footer: NodeFooter
  ): BlockPtr[] {
    const array: BlockPtr[] = [];
    for (let x = 0; x < footer.itemCount; x++) {
      const top = offset + 24 * x;
      const blockId = readNumber64(view, top + 0);
      array.push(
        {
          blockId,
          offset: readLong(view, top + 8),
          size: view.getUint16(top + 16, true),
          isData: (blockId & 2) === 0,
        } as BlockPtr
      );
    }
    return array;
  }

  export function readTablePtr(
    view: DataView,
    offset: number,
    footer: NodeFooter
  ): TablePtr[] {
    const array: TablePtr[] = [];
    for (let x = 0; x < footer.itemCount; x++) {
      const top = offset + 24 * x;
      array.push(
        {
          start: readLong(view, top + 0),
          u1: readLong(view, top + 8),
          offset: readLong(view, top + 16),
        } as TablePtr
      );
    }
    return array;
  }

  export function readNodePtr(
    view: DataView,
    offset: number,
    footer: NodeFooter
  ): NodePtr[] {
    const array: NodePtr[] = [];
    for (let x = 0; x < footer.itemCount; x++) {
      const top = offset + 32 * x;
      array.push(
        {
          nodeId: readNumber64(view, top + 0),
          blockId: readNumber64(view, top + 8),
          subBlockId: readNumber64(view, top + 16),
          parentNodeId: view.getUint32(top + 24, true),
        } as NodePtr
      );
    }
    return array;
  }

  export function readSLBlock(
    view: DataView,
    offset: number
  ): SLBlock {
    const count = view.getUint16(offset + 2, true);
    const entries: SLBlockEntry[] = [];
    for (let x = 0; x < count; x++) {
      const top = offset + 8 + 24 * x;
      entries.push(
        {
          nodeId: view.getUint32(top + 0, true),
          blockId: readNumber64(view, top + 8),
          subBlockId: readNumber64(view, top + 16),
        } as SLBlockEntry
      );
    }
    return {
      entries
    };
  }

  export function readXBlock(
    view: DataView,
    offset: number,
    itemCount: number
  ): XBlockPtr[] {
    const entries: XBlockPtr[] = [];
    for (let x = 0; x < itemCount; x++) {
      const top = offset + 8 + 8 * x;
      entries.push(
        {
          blockId: readNumber64(view, top + 0),
        } as XBlockPtr
      );
    }
    return entries;
  }

  export function readXXBlock(
    view: DataView,
    offset: number,
    itemCount: number
  ): XXBlockPtr[] {
    const entries: XXBlockPtr[] = [];
    for (let x = 0; x < itemCount; x++) {
      const top = offset + 8 + 8 * x;
      entries.push(
        {
          blockId: readNumber64(view, top + 0),
        } as XXBlockPtr
      );
    }
    return entries;
  }

  export function readSIBlock(
    view: DataView,
    offset: number
  ): SIBlock {
    const count = view.getUint16(offset + 2, true);
    const entries: SIBlockEntry[] = [];
    for (let x = 0; x < count; x++) {
      const top = offset + 8 + 24 * x;
      entries.push(
        {
          nodeId: view.getUint32(top + 0, true),
          blockId: readNumber64(view, top + 8),
        } as SIBlockEntry
      );
    }
    return {
      entries
    };
  }
}

const ver0x0e: StoreTrait = {
  BACKLINK_OFFSET: 0x1f8,
  LEVEL_INDICATOR_OFFSET: 0x1f3,
  ITEM_COUNT_OFFSET: 0x1f0,

  ENC_OFFSET: 0x1cd,
  SECOND_POINTER_COUNT: 0xB8,
  SECOND_POINTER: 0xBC,
  INDEX_POINTER_COUNT: 0xC0,
  INDEX_POINTER: 0xC4,
  BlockSize: 512,

  readId: (view, offset) => new Long(view.getUint32(offset, true)),

  readBlockPtr: ptr32.readBlockPtr,
  readTablePtr: ptr32.readTablePtr,
  readNodePtr: ptr32.readNodePtr,
  readSLBlock: ptr32.readSLBlock,
  readSIBlock: ptr32.readSIBlock,
  readXBlock: ptr32.readXBlock,
  readXXBlock: ptr32.readXXBlock,
  unzipHook: passThru1,
  is4K: false,
  storeSupportMask: undefined,
};

const ver0x17: StoreTrait = {
  BACKLINK_OFFSET: 0x1f8,
  LEVEL_INDICATOR_OFFSET: 0x1eb,
  ITEM_COUNT_OFFSET: 0x1e8,

  ENC_OFFSET: 0x201,
  SECOND_POINTER_COUNT: 0xD8,
  SECOND_POINTER: 0xE0,
  INDEX_POINTER_COUNT: 0xE8,
  INDEX_POINTER: 0xF0,
  BlockSize: 512,

  readId: readLong,

  readBlockPtr: ptr64.readBlockPtr,
  readTablePtr: ptr64.readTablePtr,
  readNodePtr: ptr64.readNodePtr,
  readSLBlock: ptr64.readSLBlock,
  readSIBlock: ptr64.readSIBlock,
  readXBlock: ptr64.readXBlock,
  readXXBlock: ptr64.readXXBlock,
  unzipHook: passThru1,
  is4K: false,
  /**
   * STORE_ENTRYID_UNIQUE | STORE_MODIFY_OK | STORE_CREATE_OK | STORE_ATTACH_OK | STORE_OLE_OK | STORE_MV_PROPS_OK | STORE_CATEGORIZE_OK | STORE_RTF_OK | STORE_UNICODE_OK
   */
  storeSupportMask: 0x40E79,
};

const ver0x24: StoreTrait = {
  BACKLINK_OFFSET: 0xff0,
  LEVEL_INDICATOR_OFFSET: 0xfdd,
  ITEM_COUNT_OFFSET: 0xfd8,

  ENC_OFFSET: 0x201,
  SECOND_POINTER_COUNT: 0xD8,
  SECOND_POINTER: 0xE0,
  INDEX_POINTER_COUNT: 0xE8,
  INDEX_POINTER: 0xF0,
  BlockSize: 4096,

  readId: readLong,

  readBlockPtr: ptr64.readBlockPtr,
  readTablePtr: ptr64.readTablePtr,
  readNodePtr: ptr64.readNodePtr,
  readSLBlock: ptr64.readSLBlock,
  readSIBlock: ptr64.readSIBlock,
  readXBlock: ptr64.readXBlock,
  readXXBlock: ptr64.readXXBlock,
  unzipHook: willUnzip1,
  is4K: true,
  /**
   * STORE_ENTRYID_UNIQUE | STORE_MODIFY_OK | STORE_CREATE_OK | STORE_ATTACH_OK | STORE_OLE_OK | STORE_MV_PROPS_OK | STORE_CATEGORIZE_OK | STORE_RTF_OK | STORE_UNICODE_OK
   */
  storeSupportMask: 0x40E79,
};

interface SIBlockEntry {
  nodeId: number;
  blockId: number;
}

interface SIBlock {
  entries: SIBlockEntry[];
}

interface SLBlockEntry {
  nodeId: number;
  blockId: number;
  subBlockId: number;
}

interface SLBlock {
  entries: SLBlockEntry[];
}

interface NodePtr {
  nodeId: number;
  blockId: number;
  subBlockId: number;
  parentNodeId: number;
}

interface TablePtr {
  start: Long;
  u1: Long;
  offset: Long;
}

interface BlockPtr {
  blockId: number;
  offset: Long;
  size: number;

  isData: boolean;
}

interface SubNodePair {
  dataBlockId: number;
  subBlockId: number;
}

export async function openLowPst(api: ReadFileApi): Promise<PLStore> {
  const surelyRead = surelyReader(api.readFile);

  const buffer = new ArrayBuffer(1024);
  const view = new DataView(buffer, 0, 1024);
  await surelyRead(buffer, 0, 1024, 0);

  const key = '!BDN'
  if (false
    || view.getUint8(0) !== key.charCodeAt(0)
    || view.getUint8(1) !== key.charCodeAt(1)
    || view.getUint8(2) !== key.charCodeAt(2)
    || view.getUint8(3) !== key.charCodeAt(3)
  ) {
    throw new Error(
      'PSTFile::open Invalid file header (expected: "!BDN"): ' + buffer.slice(0, 4)
    )
  }

  const version = view.getUint8(10);
  let trait: StoreTrait;
  if (false) { }
  else if (version === 0x0e) {
    trait = ver0x0e;
  }
  else if (version === 0x17) {
    trait = ver0x17;
  }
  else if (version === 0x24) {
    trait = ver0x24;
  }
  else {
    throw new Error(
      'PSTFile::open Unrecognised PST File version: ' + version
    )
  }

  const encryptionType = view.getUint8(trait.ENC_OFFSET);

  if (encryptionType === 0x02) {
    throw new Error('PSTFile::open PST is encrypted')
  }

  const unzipHook = trait.unzipHook;

  async function pst_read_block_size(
    position: Long,
    size: number,
    decrypt: boolean
  ): Promise<ArrayBuffer> {
    const buffer = new ArrayBuffer(size);
    await surelyRead(buffer, 0, size, position.toNumber());

    if (decrypt) {
      if (false) { }
      else if (encryptionType === 0) {
        // plain
      }
      else if (encryptionType === 1) {
        PSTUtil.decodeArray(new Uint8Array(buffer, 0, size));
      }
      else {
        throw new Error(`Unknown encryptionType ${encryptionType}`);
      }
    }

    return await unzipHook(buffer);
  }

  async function readBlock(block: BlockPtr, decrypt: boolean): Promise<ArrayBuffer> {
    return await unzipHook(
      await pst_read_block_size(
        block.offset,
        block.size,
        block.isData && decrypt
      )
    );
  }

  const blockMap = new Map<number, BlockPtr>;
  const nodeMap = new Map<number, NodePtr>;

  async function loadBlockTree(offset: Long, linku1: Long, start_val: Long): Promise<void> {
    const buf = await pst_read_block_size(offset, trait.BlockSize, false);
    const view = new DataView(buf);

    const footer = {
      itemCount: view.getUint8(trait.ITEM_COUNT_OFFSET),
      level: view.getUint8(trait.LEVEL_INDICATOR_OFFSET),
      thisId: trait.readId(view, trait.BACKLINK_OFFSET),
    } as NodeFooter;

    if (footer.thisId.neq(linku1)) {
      throw new Error(`block buffer seems to be broken. footer id is ${footer.thisId.toNumber()}, while ${linku1.toNumber()} is expected`);
    }

    if (footer.level === 0) {
      const array = trait.readBlockPtr(view, 0, footer);
      for (let x = 0; x < footer.itemCount; x++) {
        if (x === 0 && start_val.neq(0) && start_val.neq(array[x].blockId)) {
          throw new Error(`block buffer seems to be broken. first blockId is ${array[x].blockId}, while ${start_val.toNumber()} is expected`);
        }
        if (array[x].blockId === 0) {
          throw new Error(`blockId must not be 0`);
        }
        blockMap.set(
          array[x].blockId & ~1,
          array[x]
        );
      }
    }
    else {
      const array = trait.readTablePtr(view, 0, footer);
      for (let x = 0; x < footer.itemCount; x++) {
        if (x === 0 && start_val.neq(0) && start_val.neq(array[x].start)) {
          throw new Error(`block buffer seems to be broken. first table id is ${array[x].start.toNumber()}, while ${start_val.toNumber()} is expected`);
        }
        if (array[x].start.eq(0)) {
          throw new Error(`table id must not be 0`);
        }
        await loadBlockTree(
          array[x].offset,
          array[x].u1,
          array[x].start
        );
      }
    }
  }

  const block_btree_count = trait.readId(view, trait.INDEX_POINTER_COUNT);
  const block_btree = trait.readId(view, trait.INDEX_POINTER);

  await loadBlockTree(block_btree, block_btree_count, Long.ZERO);

  async function loadNodeTree(offset: Long, linku1: Long, start_val: number, prevLevel: number): Promise<void> {
    const buf = await pst_read_block_size(offset, trait.BlockSize, false);
    const view = new DataView(buf);

    const footer = {
      itemCount: view.getUint8(trait.ITEM_COUNT_OFFSET),
      level: view.getUint8(trait.LEVEL_INDICATOR_OFFSET),
      thisId: trait.readId(view, trait.BACKLINK_OFFSET),
    } as NodeFooter;

    if (prevLevel !== Infinity) {
      if (footer.level !== prevLevel - 1) {
        throw new Error(`node level descending mismatched`);
      }
    }

    if (footer.thisId.neq(linku1)) {
      throw new Error(`node buffer seems to be broken. footer id is ${footer.thisId.toNumber()}, while ${linku1.toNumber()} is expected`);
    }

    if (footer.level === 0) {
      const array = trait.readNodePtr(view, 0, footer);
      for (let x = 0; x < footer.itemCount; x++) {
        if (x === 0 && start_val !== (0) && start_val !== (array[x].nodeId)) {
          throw new Error(`node buffer seems to be broken. first nodeId is ${array[x].nodeId}, while ${start_val} is expected`);
        }
        if (array[x].nodeId === 0) {
          break;
        }
        nodeMap.set(
          array[x].nodeId,
          array[x]
        );
      }
    }
    else {
      const array = trait.readTablePtr(view, 0, footer);
      for (let x = 0; x < footer.itemCount; x++) {
        if (x === 0 && start_val !== (0) && array[x].start.neq(start_val)) {
          throw new Error(`node buffer seems to be broken. table id is ${array[x].start.toNumber()}, while ${start_val} is expected`);
        }
        if (array[x].start.isZero()) {
          throw new Error(`table id must not be 0`);
        }
        await loadNodeTree(
          array[x].offset,
          array[x].u1,
          array[x].start.toNumber(),
          footer.level
        );
      }
    }
  }

  const node_btree_count = trait.readId(view, trait.SECOND_POINTER_COUNT);
  const node_btree = trait.readId(view, trait.SECOND_POINTER);

  await loadNodeTree(node_btree, node_btree_count, 0x21, Infinity);

  // Index children once so getChildren() below is O(children) instead of
  // scanning the whole nodeMap on every call.
  const childrenByParent = new Map<number, NodePtr[]>();
  for (const node of nodeMap.values()) {
    let siblings = childrenByParent.get(node.parentNodeId);
    if (siblings === undefined) {
      siblings = [];
      childrenByParent.set(node.parentNodeId, siblings);
    }
    siblings.push(node);
  }

  async function loadMainBlockTo(
    blockId: number,
    consumer: (block: BlockPtr) => Promise<void>
  ): Promise<void> {
    if (blockId === 0) {
      return;
    }
    const block = blockMap.get(blockId & ~1);
    if (block === undefined) {
      throw new Error(`blockId=${blockId} not found`);
    }
    if (block.isData) {
      await consumer(block);
    }
    else {
      const buf = await readBlock(block, true);
      const view = new DataView(buf);

      const bType = view.getUint8(0);
      if (bType !== 1) {
        throw new Error("btype != 1");
      }

      const level = view.getUint8(1);
      const itemCount = view.getUint16(2, true);

      if (false) { }
      else if (level === 1) {
        //XBLOCK
        const entries = trait.readXBlock(view, 0, itemCount);
        for (let x = 0; x < entries.length; x++) {
          await loadMainBlockTo(
            entries[x].blockId,
            consumer
          );
        }
      }
      else if (level === 2) {
        //XXBLOCK
        const entries = trait.readXXBlock(view, 0, itemCount);
        for (let x = 0; x < entries.length; x++) {
          await loadMainBlockTo(
            entries[x].blockId,
            consumer
          );
        }
      }
      else {
        throw new Error(`invalid level ${level}`);
      }
    }
  }

  async function readSubNode(
    blockId: number,
    subNodeMap: Map<number, SubNodePair>
  ): Promise<void> {
    if (blockId === 0) {
      return;
    }

    const block = blockMap.get(blockId);
    if (block === undefined) {
      throw new Error(`blockId ${blockId} not found`);
    }

    const buf = await readBlock(block, true);
    const view = new DataView(buf);

    if (view.getUint8(0) !== 2) {
      throw new Error("btype != 2");
    }

    const level = view.getUint8(1);

    if (level === 0) {
      //SLBLOCK
      const { entries } = trait.readSLBlock(view, 0);
      for (let index = 0; index < entries.length; index++) {
        subNodeMap.set(
          entries[index].nodeId,
          {
            dataBlockId: entries[index].blockId,
            subBlockId: entries[index].subBlockId,
          } as SubNodePair
        );
      }
    }
    else {
      //SIBLOCK
      const struc = trait.readSIBlock(view, 0);
      const { entries } = struc;
      for (let index = 0; index < entries.length; index++) {
        await readSubNode(
          entries[index].blockId,
          subNodeMap
        );
      }
    }
  }

  //const nodeMap = processNameToIDMap();

  function getOneNodeBy(nodeId: number): PLNode | undefined {
    if (nodeId === 0) {
      return undefined;
    }

    const ptr = nodeMap.get(nodeId);
    if (ptr === undefined) {
      return undefined;
    }

    const { blockId } = ptr;

    async function getData(): Promise<ArrayBuffer> {
      const array: ArrayBuffer[] = [];

      await loadMainBlockTo(
        blockId,
        async (block): Promise<void> => {
          const buf = await readBlock(block, true);
          array.push(buf);
        }
      );

      if (array.length === 0) {
        return new ArrayBuffer(0);
      }
      else {
        return array[0];
      }
    }

    async function getDataOf(blockId: number): Promise<ArrayBuffer[]> {
      const array = [] as ArrayBuffer[];
      await loadMainBlockTo(
        blockId,
        async (block) => {
          array.push(await readBlock(block, true));
        }
      );
      return array;
    }

    async function getChildOf(
      blockId: number,
      childNodeId: number,
      parentToString: string
    ): Promise<PLSubNode | undefined> {
      if (blockId === 0) {
        return undefined;
      }
      const block = blockMap.get(blockId & ~1);
      if (block === undefined) {
        throw new Error(
          `blockId=${blockId}`
          + ` for childNodeId=0x${childNodeId.toString(16)}`
          + ` of ${parentToString}`
          + ` not found`
        );
      }

      const subNodeMap = new Map<number, SubNodePair>();

      await readSubNode(
        block.blockId,
        subNodeMap
      );

      const thisToString =
        `childNodeId=0x${childNodeId.toString(16)}`
        + `,nidType=0x${(childNodeId & 0x1f).toString(16)}`
        + ` of ${parentToString}`;

      const subNode = subNodeMap.get(childNodeId);
      if (subNode === undefined) {
        return undefined;
      }

      return {
        nodeId: childNodeId,
        is4K: trait.is4K,
        getChildBy: async (childNodeId) => await getChildOf(
          subNode.subBlockId,
          childNodeId,
          thisToString
        ),
        getData: async () => await getDataOf(subNode.dataBlockId),
        toString: () => thisToString,
      } as PLSubNode;
    }

    function getSubNodeOf(nodeId: number): PLSubNode {
      const node = nodeMap.get(nodeId);
      if (node === undefined) {
        throw new Error(`nodeId=${nodeId} not found`);
      }

      const thisToString = `subNode of nodeId=${nodeId},nidType=${nodeId & 0x1f}`;

      return {
        nodeId: nodeId,
        is4K: trait.is4K,
        getChildBy: async (childNodeId) => getChildOf(
          node.subBlockId,
          childNodeId,
          thisToString
        ),
        getData: async () => await getDataOf(node.blockId),
        toString: () => thisToString,
      } as PLSubNode;
    }

    return {
      nodeId: (ptr.nodeId),
      getParent: () => getOneNodeBy((ptr.parentNodeId)),
      getChildren: () => (childrenByParent.get(ptr.nodeId) ?? [])
        .filter(it => it.nodeId !== ptr.nodeId)
        .map(it => getOneNodeBy(it.nodeId))
        .filter(it => it !== undefined),
      getSubNode: () => getSubNodeOf(ptr.nodeId),
      getSiblingNode: (nidType: number) => getOneNodeBy((nodeId & ~0x1f) | (nidType & 0x1f)),
    } as PLNode;
  };

  function getOneNodeByOrError(nodeId: number): PLNode {
    const node = getOneNodeBy(nodeId);
    if (node === undefined) {
      throw new Error(`node ${nodeId} must exist and must be valid`);
    }
    return node;
  }

  return {
    getOneNodeBy,
    getOneNodeByOrError,
    close: () => api.close(),
    storeSupportMask: trait.storeSupportMask,
  } as PLStore;
}
