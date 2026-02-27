import { KeyedDelay } from "./KeyedDelay.js";

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
    this._cache = new KeyedDelay<T>();
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
