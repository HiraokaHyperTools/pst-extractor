import { KeyedDelay } from "./KeyedDelay.js";

/**
 * How many items one collection keeps.
 *
 * The cache makes reading the same item twice cheap. Unbounded, it also means
 * a collection holds every item it has ever produced for as long as it is
 * alive: a folder's message collection lives on the PSTFolder, so walking a
 * large mailbox once kept every message in memory (about 1.3 GB of parsed
 * messages for a 2 GB file), which defeats reading a folder incrementally.
 *
 * @internal
 */
const MAX_CACHED_ITEMS = 64;

/**
 * @internal
 */
export class CollectionAsyncProvider<T> {
  private _cache: KeyedDelay<T>;
  private _count: number;
  private _itemProvider: (index: number) => Promise<T>;

  constructor(
    count: number,
    itemProvider: (index: number) => Promise<T>
  ) {
    this._cache = new KeyedDelay<T>(MAX_CACHED_ITEMS);
    this._count = count;
    this._itemProvider = itemProvider;
  }

  public get count() {
    return this._count;
  }

  public async get(index: number): Promise<T> {
    return await this._cache.getOrCreate(
      index.toString(),
      async () => await this._itemProvider(index)
    );
  }

  /**
   * get all of the children
   */
  public async all(): Promise<T[]> {
    const array: T[] = [];
    for (let x = 0; x < this._count; x++) {
      array.push(await this.get(x));
    }
    return array;
  }
}
