import { type BTHeapEntry, getBTHeapReaderFrom } from "./BTHeap.js";
import type { PHNodeHeap } from "./PHNodeHeap.js";
import type { Property } from "./Property.js";
import type { PropertyContext } from "./PropertyContext.js";
import type { PropertyValueResolver } from "./PropertyValueResolver.js";
import type { RawProperty } from "./RawProperty.js";

const bTypePC = 0xBC;

function copy<T extends {}>(array: T[]): T[] {
  return array
    .map(it => Object.assign({}, it));
}

export async function getPropertyContext(
  heap: PHNodeHeap,
  resolver: PropertyValueResolver
): Promise<PropertyContext> {
  if (heap.bClientSig !== bTypePC) {
    throw new Error("bClientSig must be bTypePC");
  }
  const reader = heap.getReader();

  async function getRecords(): Promise<BTHeapEntry[]> {
    try {
      const index_data = await getBTHeapReaderFrom(reader, heap.userRootHnid);
      return await index_data.list();
    }
    catch (ex) {
      throw new Error(`getRecords failure of hnid ${heap.userRootHnid} of ${heap} --> ${ex}`);
    }
  }

  const rawProperties: RawProperty[] = (await getRecords())
    .map(
      record => {
        const keyView = new DataView(record.key);
        const dataView = new DataView(record.data);

        return {
          key: keyView.getUint16(0, true),
          type: dataView.getUint16(0, true),
          value: record.data.slice(2),
        };
      }
    );

  return {
    async listRaw(): Promise<RawProperty[]> {
      return copy(rawProperties);
    },
    async list(): Promise<Property[]> {
      const props = [];
      for (let it of rawProperties) {
        try {
          props.push(
            {
              key: it.key,
              type: it.type,
              value: await resolver.resolveValueOf(it.key, it.type, it.value, reader),
            } as Property
          );
        }
        catch (ex) {
          throw new Error(`getPropertyContext.list() resolving property`
            + ` key=0x${it.key.toString(16).padStart(4, '0')}`
            + ` type=0x${it.type.toString(16).padStart(4, '0')}`
            + ` of ${heap} failure`
            + ` --> ${ex}`
          );
        }
      }
      return props;
    },
  } as PropertyContext;
}
