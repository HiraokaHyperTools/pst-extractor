/**
 * @internal
 */
export class KeyedDelay<T> {
  private _map: Map<string, T>;
  private _capacity: number;

  /**
   * @param capacity how many entries to keep. Unlimited by default, so
   * existing callers are unaffected; a caller that may produce a very large
   * number of entries can cap it and have the least recently used dropped.
   */
  constructor(capacity: number = Number.POSITIVE_INFINITY) {
    this._map = new Map<string, T>();
    this._capacity = capacity;
  }

  async getOrCreate(key: string, provider: () => Promise<T>): Promise<T> {
    if (this._map.has(key)) {
      const value = this._map.get(key);
      if (value === undefined) {
        throw new Error("provider must provide a non-undefined value");
      }
      // Re-insert so that Map iteration order tracks recency of use, which is
      // what makes the eviction below drop the coldest entry.
      this._map.delete(key);
      this._map.set(key, value);
      return value;
    }
    else {
      const value = await provider();
      this._map.set(key, value);
      while (this._map.size > this._capacity) {
        const oldest = this._map.keys().next().value;
        if (oldest === undefined) {
          break;
        }
        this._map.delete(oldest);
      }
      return value;
    }
  }
}
