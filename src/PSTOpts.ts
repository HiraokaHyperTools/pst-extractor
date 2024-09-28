import { PrimitiveTypeConverter } from "./PropertyValueResolverV1";

export interface PSTOpts {
  /**
   * Specify character encoding usable with `iconv-lite` package.
   * 
   * This is used to obtain unicode string from ArrayBuffer of ansiString PT_STRING8.
   * 
   * It is rare case that pst/ost files store ansiString,
   * because recent Outlook applications are Unicode based.
   * 
   * This applies to pst files generated by older versions of Outlook.
   * 
   * And {@link ansiEncoding} is required because ANSI code page depends on
   * Windows's system locale can be acquired by [GetOEMCP](https://docs.microsoft.com/en-us/windows/win32/api/winnls/nf-winnls-getoemcp) API.
   * 
   * OEMCP cannot be obtained from pure JavaScript environment,
   * especially on Web browser environment.
   * Thus you need to specify considered one.
   */
  ansiEncoding?: string;

  /**
   * Otherwise, you can provide your own converter (from ansiString to unicode string).
   */
  convertAnsiString?: (arrayBuffer: ArrayBuffer) => Promise<string>;

  /**
   * Provide your own optional type converter.
   * 
   * About known types, check:
   * 
   * - [[MS-OXCDATA]: Property Data Types | Microsoft Learn](https://learn.microsoft.com/ja-jp/openspecs/exchange_server_protocols/ms-oxcdata/0c77892e-288e-435a-9c49-be1c20c7afdb)
   * 
   * A type converter is resolved in this order:
   * 
   * - `provideTypeConverterOf` (called if provided. return `undefined` to go next resolver)
   * - `typeConverters` (built-in type converters)
   * - `throw new Error(...);`
   * 
   * @example
   * ```ts
   * const pst = openPstFile(
   *   'path/to/file.pst',
   *   {
   *     provideTypeConverterOf: (propertyType) => {
   *       if (propertyType === 0x1002) {
   *         return async (arg) => {
   *           const heap = arg.view.getUint32(0, true);
   *           const list = [] as any[];
   *           if (heap !== 0) {
   *             const bytes = await arg.resolveHeap(heap);
   *             if (bytes !== undefined) {
   *               const view = new DataView(bytes);
   *               const count = bytes.byteLength / 2;
   *               for (let x = 0; x < count; x++) {
   *                 list.push(view.getInt16(2 * x, true))
   *               }
   *             }
   *           }
   *           return list;
   *         };
   *       }
   *       return undefined;
   *     },
   *   }
   * );
   * 
   * ```
   * 
   * @param propertyType A numeric like `0x1002`
   * @returns A valid PrimitiveTypeConverter; otherwise,
   * return `undefined` to indicate to find another type converter.
   */
  provideTypeConverterOf?: (propertyType: number) => PrimitiveTypeConverter | undefined;
}
