/**
 * PST adapter utilities
 */

import { msftUuidStringify } from "./msftUuidStringify";
import { NodeMap } from "./NodeMap.class"
import { getHeapFrom } from "./PHUtil";
import { PLNode } from "./PLNode"
import { Property } from "./Property";
import { getPropertyContext } from "./PropertyContextUtil";
import { PropertyValueResolver } from "./PropertyValueResolver";
import { PSTUtil } from "./PSTUtil.class";

export interface PropertyFinder {
  findByKey(key: number): Property | undefined;
  all(): Property[];
}

const guidMap: Map<string, number> = new Map([
  ['00020329-0000-0000-C000-000000000046', 0],
  ['00062008-0000-0000-C000-000000000046', 1],
  ['00062004-0000-0000-C000-000000000046', 2],
  ['00020386-0000-0000-C000-000000000046', 3],
  ['00062002-0000-0000-C000-000000000046', 4],
  ['6ED8DA90-450B-101B-98DA-00AA003F1305', 5],
  ['0006200A-0000-0000-C000-000000000046', 6],
  ['41F28F13-83F4-4114-A584-EEDB5A6B0BFF', 7],
  ['0006200E-0000-0000-C000-000000000046', 8],
  ['00062041-0000-0000-C000-000000000046', 9],
  ['00062003-0000-0000-C000-000000000046', 10],
  ['4442858E-A9E3-4E80-B900-317A210CC15B', 11],
  ['00020328-0000-0000-C000-000000000046', 12],
  ['71035549-0739-4DCB-9163-00F0580DBBDF', 13],
  ['00062040-0000-0000-C000-000000000046', 14],
]);

/**
 * Process name to ID map.
 * 
 * @param nameToIdMapDescriptorNode nodeId 97
 */
export async function processNameToIDMap(
  nameToIdMapDescriptorNode: PLNode,
  resolver: PropertyValueResolver
): Promise<NodeMap> {
  const subNode = nameToIdMapDescriptorNode.getSubNode();
  const bcTable = (await (await getPropertyContext(
    await getHeapFrom(subNode),
    resolver
  ))
    .list());

  async function getTableItem(key: number): Promise<ArrayBuffer | undefined> {
    const found = bcTable.find(it => it.key === key);
    if (found === undefined) {
      throw new Error(`processNameToIDMap key:${key} is null`);
    }
    const { value } = found;
    if (value instanceof ArrayBuffer) {
      return value;
    }
    return undefined;
  }

  // process the map

  // Get the guids
  const guidEntry = await getTableItem(2);
  if (!guidEntry) {
    throw new Error('PSTFile::processNameToIDMap guidEntry is null')
  }

  const guids = new Uint8Array(guidEntry);
  const nGuids = Math.trunc(guids.byteLength / 16);
  const guidIndexes: number[] = []
  let offset = 0
  for (let i = 0; i < nGuids; ++i) {
    const strUID: string = msftUuidStringify(guids, offset).toUpperCase();

    const guid = guidMap.get(strUID)
    if (guid) {
      guidIndexes[i] = guid
    } else {
      guidIndexes[i] = -1 // We don't know this guid
    }
    // console.log('PSTFile:: processNameToIdMap idx: ' + i + ', ' + strUID + ', ' + guidIndexes[i]);
    offset += 16
  }

  // if we have a reference to an internal descriptor
  const mapEntries = await getTableItem(3);
  if (!mapEntries) {
    throw new Error('PSTFile::processNameToIDMap mapEntries is null')
  }
  const nameToIdByte: Uint8Array = new Uint8Array(mapEntries);
  const nameToIdByteView = new DataView(mapEntries);

  const stringMapEntries = await getTableItem(4);
  if (!stringMapEntries) {
    throw new Error('PSTFile::processNameToIDMap stringMapEntries is null')
  }
  const stringNameToIdByte: Uint8Array = new Uint8Array(
    stringMapEntries
  );
  const stringNameToIdByteView = new DataView(stringMapEntries);
  const stringNameToIdByteBuffer = Buffer.from(stringMapEntries);

  const nodeMap = new NodeMap();

  // process the entries
  for (let x = 0; x + 8 < nameToIdByte.length; x += 8) {
    const key: number = nameToIdByteView.getUint32(x, true);
    let guid: number = nameToIdByteView.getUint16(x + 4, true);
    let propId: number = nameToIdByteView.getUint16(x + 6, true);

    if (key == 0x55555555) {
      break;
    }

    const PS_PUBLIC_STRINGS = 0;
    const PS_MAPI = 12;

    if ((guid & 0x0001) == 0) {
      // identifier is numeric
      propId += 0x8000
      guid >>= 1
      let guidIndex: number
      if (guid == 1) {
        guidIndex = PS_MAPI
      } else if (guid == 2) {
        guidIndex = PS_PUBLIC_STRINGS
      } else {
        guidIndex = guidIndexes[guid - 3]
      }
      nodeMap.setId(key, propId, guidIndex)
    } else {
      // identifier is a string
      // key is byte offset into the String stream in which the string name of the property is stored.
      const len = stringNameToIdByteView.getUint32(key, true);
      const keyByteValue = Buffer.alloc(len)
      PSTUtil.arraycopy(
        stringNameToIdByteBuffer,
        key + 4,
        keyByteValue,
        0,
        keyByteValue.length
      )
      propId += 0x8000
      nodeMap.setId(
        keyByteValue.toString('utf16le').replace(/\0/g, ''),
        propId
      )
    }
  }

  return nodeMap;
}

export function createPropertyFinder(props: Property[]): PropertyFinder {
  return {
    findByKey(key) {
      return props.find(it => it.key === key);
    },
    all() {
      return props;
    }
  } as PropertyFinder;
}
