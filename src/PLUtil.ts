/**
 * PST lower level utilities
 */

import Long from "long";
import { PLNode } from "./PLNode";
import { PLStore } from "./PLStore";
import { PLSubNode } from "./PLSubNode";
import { PSTUtil } from "./PSTUtil.class";
import * as zlib from 'zlib'

export type ReadFile = (buffer: ArrayBuffer, offset: number, length: number, position: number) => Promise<number>;

export interface ReadFileApi {
  readFile: ReadFile,
  close: () => Promise<void>,
}

function surelyReader(readFile: ReadFile): ReadFile {
  return async (buffer: ArrayBuffer, offset: number, length: number, position: number): Promise<number> => {
    const bytesRead = await readFile(buffer, offset, length, position);
    if (bytesRead !== length) {
      throw new Error("EOS");
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
  ReadId(view: DataView, offset: number): Long;

  ENC_OFFSET: number;
  SECOND_POINTER_COUNT: number;
  SECOND_POINTER: number;
  INDEX_POINTER_COUNT: number;
  INDEX_POINTER: number;
  BlockSize: number;
  unzipHook: UnzipHook;
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
      const arrayBuffer = await new Promise<ArrayBuffer>(
        (resolve, reject) => zlib.unzip(
          data,
          (error, result) => {
            if (error) {
              reject(error);
            }
            else {
              resolve(
                result.buffer.slice(
                  result.byteOffset,
                  result.byteLength
                )
              );
            }
          }
        )
      );
      return arrayBuffer;
    }
  }

  return data;
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

  ReadId: readLong,

  readBlockPtr: ptr64.readBlockPtr,
  readTablePtr: ptr64.readTablePtr,
  readNodePtr: ptr64.readNodePtr,
  readSLBlock: ptr64.readSLBlock,
  readSIBlock: ptr64.readSIBlock,
  readXBlock: ptr64.readXBlock,
  readXXBlock: ptr64.readXXBlock,
  unzipHook: passThru1,
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

  ReadId: readLong,

  readBlockPtr: ptr64.readBlockPtr,
  readTablePtr: ptr64.readTablePtr,
  readNodePtr: ptr64.readNodePtr,
  readSLBlock: ptr64.readSLBlock,
  readSIBlock: ptr64.readSIBlock,
  readXBlock: ptr64.readXBlock,
  readXXBlock: ptr64.readXXBlock,
  unzipHook: willUnzip1,
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
      thisId: trait.ReadId(view, trait.BACKLINK_OFFSET),
    } as NodeFooter;

    if (footer.thisId.neq(linku1)) {
      throw new Error("blah 1");
    }

    if (footer.level === 0) {
      const array = trait.readBlockPtr(view, 0, footer);
      for (let x = 0; x < footer.itemCount; x++) {
        if (x === 0 && start_val.neq(0) && start_val.neq(array[x].blockId)) {
          throw new Error("blah 3");
        }
        if (array[x].blockId === 0) {
          throw new Error("OHNO");
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
          throw new Error("blah 3");
        }
        if (array[x].start.eq(0)) {
          throw new Error("OHNO");
        }
        await loadBlockTree(
          array[x].offset,
          array[x].u1,
          array[x].start
        );
      }
    }
  }

  const block_btree_count = trait.ReadId(view, trait.INDEX_POINTER_COUNT);
  const block_btree = trait.ReadId(view, trait.INDEX_POINTER);

  await loadBlockTree(block_btree, block_btree_count, Long.ZERO);

  async function loadNodeTree(offset: Long, linku1: Long, start_val: number, prevLevel: number): Promise<void> {
    const buf = await pst_read_block_size(offset, trait.BlockSize, false);
    const view = new DataView(buf);

    const footer = {
      itemCount: view.getUint8(trait.ITEM_COUNT_OFFSET),
      level: view.getUint8(trait.LEVEL_INDICATOR_OFFSET),
      thisId: trait.ReadId(view, trait.BACKLINK_OFFSET),
    } as NodeFooter;

    if (prevLevel !== Infinity) {
      if (footer.level !== prevLevel - 1) {
        throw new Error("ohno");
      }
    }

    if (footer.thisId.neq(linku1)) {
      throw new Error("blah1");
    }

    if (footer.level === 0) {
      const array = trait.readNodePtr(view, 0, footer);
      for (let x = 0; x < footer.itemCount; x++) {
        if (x === 0 && start_val !== (0) && start_val !== (array[x].nodeId)) {
          throw new Error("blah 3");
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
          throw new Error("blah 3");
        }
        if (array[x].start.isZero()) {
          throw new Error("OHNO");
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

  const node_btree_count = trait.ReadId(view, trait.SECOND_POINTER_COUNT);
  const node_btree = trait.ReadId(view, trait.SECOND_POINTER);

  await loadNodeTree(node_btree, node_btree_count, 0x21, Infinity);

  async function loadMainBlockTo(
    blockId: number,
    consumer: (block: BlockPtr) => Promise<void>
  ): Promise<void> {
    if (blockId === 0) {
      return;
    }
    const block = blockMap.get(blockId);
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
        throw new Error(`Invalid level ${level}`);
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
    ): Promise<PLSubNode> {
      const block = blockMap.get(blockId);
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

      const subNode = subNodeMap.get(childNodeId);
      if (subNode === undefined) {
        throw new Error(
          `childNodeId=0x${childNodeId.toString(16)}`
          + `,nidType=0x${(childNodeId & 0x1f).toString(16)}`
          + ` not found`
        );
      }

      const thisToString =
        `childNodeId=0x${childNodeId.toString(16)}`
        + `,nidType=0x${(childNodeId & 0x1f).toString(16)}`
        + ` of ${parentToString}`;

      return {
        nodeId: childNodeId,
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
      getChildren: () => Array.from(nodeMap.values())
        .filter(it => it.parentNodeId === ptr.nodeId && it.nodeId !== ptr.nodeId)
        .map(it => getOneNodeBy(it.nodeId))
        .filter(it => it !== undefined),
      getSubNode: () => getSubNodeOf(ptr.nodeId),
      getSiblingNode: (nidType: number) => getOneNodeBy((nodeId & ~0x1f) | (nidType & 0x1f)),
    } as PLNode;
  };

  function getOneNodeByOrError(nodeId: number): PLNode {
    const node = getOneNodeBy(nodeId);
    if (node === undefined) {
      throw new Error(`node ${nodeId} must be valid`);
    }
    return node;
  }

  return {
    getOneNodeBy,
    getOneNodeByOrError,
    close: () => api.close(),
  } as PLStore;
}
