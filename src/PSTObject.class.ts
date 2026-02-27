/* eslint-disable @typescript-eslint/no-explicit-any */
import Long from 'long';
import { OutlookProperties } from './OutlookProperties.js';
import type { PropertyFinder } from './PAUtil.js';
import type { PLNode } from './PLNode.js';
import type { PLSubNode } from './PLSubNode.js';
import type { Property } from './Property.js';
import { PSTFile } from './PSTFile.class.js';
import type { RootProvider } from './RootProvider.js';

export abstract class PSTObject {
  protected _rootProvider: RootProvider
  protected _node: PLNode
  protected _subNode: PLSubNode
  protected _propertyFinder: PropertyFinder

  /**
   * Creates an instance of PSTObject, the root class of most PST Items.
   * @internal
   * @memberof PSTObject
   */
  constructor(
    rootProvider: RootProvider,
    node: PLNode,
    subNode: PLSubNode,
    propertyFinder: PropertyFinder
  ) {
    if (!propertyFinder) {
      console.trace("propertyFinder not defined");
    }
    this._rootProvider = rootProvider
    this._node = node;
    this._subNode = subNode;
    this._propertyFinder = propertyFinder;
  }

  /**
   * Get the node type for the descriptor id.
   * @param {number} [descriptorIdentifier]
   * @returns {number}
   * @memberof PSTObject
   */
  protected getNodeType(descriptorIdentifier?: number): number {
    if (descriptorIdentifier) {
      return descriptorIdentifier & 0x1f
    } else if (this._node.nodeId) {
      return this._node.nodeId & 0x1f
    } else {
      return -1
    }
  }

  /**
   * @protected
   * @param { number } identifier
   * @param { number } [defaultValue]
   * @returns { number }
   * @memberof PSTObject
   */
  protected getIntItem(identifier: number, defaultValue?: number): number {
    if (!defaultValue) {
      defaultValue = 0
    }
    const property = this._propertyFinder.findByKey(identifier);
    if (property !== undefined) {
      const { value } = property;
      if (typeof value === 'number') {
        return value;
      }
    }
    return defaultValue
  }

  /**
   * Get a boolean.
   * @protected
   * @param {number} identifier
   * @param {boolean} [defaultValue]
   * @returns {boolean}
   * @memberof PSTObject
   */
  protected getBooleanItem(
    identifier: number,
    defaultValue?: boolean
  ): boolean {
    if (defaultValue === undefined) {
      defaultValue = false
    }
    const property = this._propertyFinder.findByKey(identifier);
    if (property !== undefined) {
      const { value } = property;
      if (typeof value === 'boolean') {
        return value;
      }
    }
    return defaultValue
  }

  /**
   * Get a double.
   * @protected
   * @param {number} identifier
   * @param {number} [defaultValue]
   * @returns {number}
   * @memberof PSTObject
   */
  protected getDoubleItem(identifier: number, defaultValue?: number): number {
    if (defaultValue === undefined) {
      defaultValue = 0
    }
    const property = this._propertyFinder.findByKey(identifier);
    if (property !== undefined) {
      const { value } = property;
      if (typeof value === 'number') {
        return value;
      }
    }
    return defaultValue
  }

  /**
   * Get a long.
   * @protected
   * @param {number} identifier
   * @param {long} [defaultValue]
   * @returns {long}
   * @memberof PSTObject
   */
  protected getLongItem(identifier: number, defaultValue?: Long): Long {
    if (defaultValue === undefined) {
      defaultValue = Long.ZERO
    }
    const property = this._propertyFinder.findByKey(identifier);
    if (property !== undefined) {
      const { value } = property;
      if (value instanceof Long) {
        return value;
      }
      else if (typeof value === 'number') {
        return new Long(value);
      }
    }
    return defaultValue
  }

  /**
   * Get a string.
   * @protected
   * @param {number} identifier
   * @param {number} [stringType]
   * @param {string} [codepage]
   * @returns {string}
   * @memberof PSTObject
   */
  protected getStringItem(
    identifier: number,
    stringType?: number,
    codepage?: string
  ): string {
    const property = this._propertyFinder.findByKey(identifier);
    if (property !== undefined) {
      const { value } = property;
      if (typeof value === 'string') {
        return value
      }
    }
    return ''
  }

  /**
   * Get a date.
   * @param {number} identifier
   * @returns {Date}
   * @memberof PSTObject
   */
  protected getDateItem(identifier: number): Date | null {
    const property = this._propertyFinder.findByKey(identifier);
    if (property !== undefined) {
      const { value } = property;
      if (value instanceof Date) {
        return value;
      }
    }
    return null
  }

  /**
   * Get a blob.
   * @protected
   * @param {number} identifier
   * @returns {Buffer}
   * @memberof PSTObject
   */
  protected getBinaryItem(identifier: number): Buffer | null {
    const property = this._propertyFinder.findByKey(identifier);
    if (property !== undefined) {
      const { value } = property;
      if (value instanceof ArrayBuffer) {
        return Buffer.from(value);
      }
    }
    return null
  }

  /**
   * Get the display name of this object.
   * https://msdn.microsoft.com/en-us/library/office/cc842383.aspx
   * @readonly
   * @type {string}
   * @memberof PSTObject
   */
  public get displayName(): string {
    return this.getStringItem(OutlookProperties.PR_DISPLAY_NAME)
  }

  /**
   * Try to get specified property from PropertyContext.
   * 
   * @param key `0x3001` is `PR_DISPLAY_NAME` for example
   * @returns The found one will be returned. Otherwise `undefined` is returned.
   */
  getProperty(key: number): Property | undefined {
    return this._propertyFinder.findByKey(key);
  }

  /**
   * Get all properties.
   * 
   * @returns All properties.
   */
  getAllProperties(): Property[] {
    return this._propertyFinder.all();
  }

  /**
   * Get the primary node ID.
   */
  public get primaryNodeId(): number {
    return this._node.nodeId;
  }

  /**
   * JSON the object.
   * @returns {string}
   * @memberof PSTObject
   */
  public toJSON(): any {
    return this
  }
}
