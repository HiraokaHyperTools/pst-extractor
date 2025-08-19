/* eslint-disable @typescript-eslint/no-explicit-any */
import { PSTFile } from './PSTFile.class'
import { PSTMessage } from './PSTMessage.class'
import { OutlookProperties } from './OutlookProperties'
import { PLNode } from './PLNode'
import { PropertyFinder } from './PAUtil'
import { PLSubNode } from './PLSubNode'
import { RootProvider } from './RootProvider'

export class PSTActivity extends PSTMessage {
  /**
   * Creates an instance of PSTActivity.  Represents Journal entries, class IPM.Activity.
   * https://msdn.microsoft.com/en-us/library/office/aa204771(v=office.11).aspx
   * @internal
   * @param {PSTFile} rootProvider
   * @param {DescriptorIndexNode} descriptorIndexNode
   * @param {Map<number, PSTDescriptorItem>} [localDescriptorItems]
   * @memberof PSTActivity
   */
  constructor(
    rootProvider: RootProvider,
    node: PLNode,
    subNode: PLSubNode,
    propertyFinder: PropertyFinder
  ) {
    super(rootProvider, node, subNode, propertyFinder);
  }

  /**
   * Contains the display name of the journaling application (for example, "MSWord"), and is typically a free-form attribute of a journal message, usually a string.
   * https://msdn.microsoft.com/en-us/library/office/cc839662.aspx
   * @readonly
   * @type {string}
   * @memberof PSTActivity
   */
  public get logType(): string {
    return this.getStringItem(
      this._rootProvider.getNameToIdMapItem(
        OutlookProperties.PidLidLogType,
        OutlookProperties.PSETID_Log
      )
    )
  }

  /**
   * Represents the start date and time for the journal message.
   * https://msdn.microsoft.com/en-us/library/office/cc842339.aspx
   * @readonly
   * @type {Date}
   * @memberof PSTActivity
   */
  public get logStart(): Date | null {
    return this.getDateItem(
      this._rootProvider.getNameToIdMapItem(
        OutlookProperties.PidLidLogStart,
        OutlookProperties.PSETID_Log
      )
    )
  }

  /**
   * Represents the duration, in minutes, of a journal message.
   * https://msdn.microsoft.com/en-us/library/office/cc765536.aspx
   * @readonly
   * @type {number}
   * @memberof PSTActivity
   */
  public get logDuration(): number {
    return this.getIntItem(
      this._rootProvider.getNameToIdMapItem(
        OutlookProperties.PidLidLogDuration,
        OutlookProperties.PSETID_Log
      )
    )
  }

  /**
   * Represents the end date and time for the journal message.
   * https://msdn.microsoft.com/en-us/library/office/cc839572.aspx
   * @readonly
   * @type {Date}
   * @memberof PSTActivity
   */
  public get logEnd(): Date | null {
    return this.getDateItem(
      this._rootProvider.getNameToIdMapItem(
        OutlookProperties.PidLidLogEnd,
        OutlookProperties.PSETID_Log
      )
    )
  }

  /**
   * Contains metadata about the journal.
   * https://msdn.microsoft.com/en-us/library/office/cc815433.aspx
   * @readonly
   * @type {number}
   * @memberof PSTActivity
   */
  public get logFlags(): number {
    return this.getIntItem(
      this._rootProvider.getNameToIdMapItem(
        OutlookProperties.PidLidLogFlags,
        OutlookProperties.PSETID_Log
      )
    )
  }

  /**
   * Indicates whether the document was printed during journaling.
   * https://msdn.microsoft.com/en-us/library/office/cc839873.aspx
   * @readonly
   * @type {boolean}
   * @memberof PSTActivity
   */
  public get isDocumentPrinted(): boolean {
    return this.getBooleanItem(
      this._rootProvider.getNameToIdMapItem(
        OutlookProperties.PidLidLogDocumentPrinted,
        OutlookProperties.PSETID_Log
      )
    )
  }

  /**
   * Indicates whether the document was saved during journaling.
   * https://msdn.microsoft.com/en-us/library/office/cc815488.aspx
   * @readonly
   * @type {boolean}
   * @memberof PSTActivity
   */
  public get isDocumentSaved(): boolean {
    return this.getBooleanItem(
      this._rootProvider.getNameToIdMapItem(
        OutlookProperties.PidLidLogDocumentSaved,
        OutlookProperties.PSETID_Log
      )
    )
  }

  /**
   * Indicates whether the document was sent to a routing recipient during journaling.
   * https://msdn.microsoft.com/en-us/library/office/cc839558.aspx
   * @readonly
   * @type {boolean}
   * @memberof PSTActivity
   */
  public get isDocumentRouted(): boolean {
    return this.getBooleanItem(
      this._rootProvider.getNameToIdMapItem(
        OutlookProperties.PidLidLogDocumentRouted,
        OutlookProperties.PSETID_Log
      )
    )
  }

  /**
   * Indicates whether the document was sent by e-mail or posted to a server folder during journaling.
   * https://msdn.microsoft.com/en-us/library/office/cc815353.aspx
   * @readonly
   * @type {boolean}
   * @memberof PSTActivity
   */
  public get isDocumentPosted(): boolean {
    return this.getBooleanItem(
      this._rootProvider.getNameToIdMapItem(
        OutlookProperties.PidLidLogDocumentPosted,
        OutlookProperties.PSETID_Log
      )
    )
  }

  /**
   * Describes the activity that is being recorded.
   * https://msdn.microsoft.com/en-us/library/office/cc815500.aspx
   * @readonly
   * @type {string}
   * @memberof PSTActivity
   */
  public get logTypeDesc(): string {
    return this.getStringItem(
      this._rootProvider.getNameToIdMapItem(
        OutlookProperties.PidLidLogTypeDesc,
        OutlookProperties.PSETID_Log
      )
    )
  }

  /**
   * JSON stringify the object properties.
   * @returns {string}
   * @memberof PSTActivity
   */
  public toJSON(): any {
    const clone = Object.assign(
      {
        messageClass: this.messageClass,
        subject: this.subject,
        importance: this.importance,
        transportMessageHeaders: this.transportMessageHeaders,
        logType: this.logType,
        logStart: this.logStart,
        logDuration: this.logDuration,
        logEnd: this.logEnd,
        logFlags: this.logFlags,
        isDocumentPrinted: this.isDocumentPrinted,
        isDocumentSaved: this.isDocumentSaved,
        isDocumentRouted: this.isDocumentRouted,
        isDocumentPosted: this.isDocumentPosted,
        logTypeDesc: this.logTypeDesc,
        primaryNodeId: this.primaryNodeId,
      },
      this
    )
    return clone
  }
}
