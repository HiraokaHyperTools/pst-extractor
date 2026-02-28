
/**
 * The value representation of `PT_OBJECT` type
 * 
 * @see {@link https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-pst/49457d57-820e-453d-bbc0-1d192a999814 | [MS-PST]: PtypObject Properties | Microsoft Docs}
 */
export class PropertyTypeObject {
  private _subNodeId: number;
  private _size: number;

  constructor(subNodeId: number, size: number) {
    this._subNodeId = subNodeId;
    this._size = size;
  }

  get subNodeId(): number {
    return this._subNodeId;
  }

  get size(): number {
    return this._size;
  }

}
