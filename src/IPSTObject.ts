import type { Property } from "./Property.js";

export interface IPSTObject {
  /**
   * Get the display name of this object.
   * https://msdn.microsoft.com/en-us/library/office/cc842383.aspx
   * @readonly
   */
  get displayName(): string;

  /**
   * Try to get specified property from PropertyContext.
   * 
   * @param key `0x3001` is `PR_DISPLAY_NAME` for example
   * @returns The found one will be returned. Otherwise `undefined` is returned.
   */
  getProperty(key: number): Property | undefined;

  /**
   * Get all properties.
   * 
   * @returns All properties.
   */
  getAllProperties(): Property[];

  /**
   * Get the primary node ID.
   */
  get primaryNodeId(): number;

  /**
   * JSON the object.
   */
  toJSON(): any;
}
