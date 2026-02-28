/* eslint-disable @typescript-eslint/no-explicit-any */
import { OutlookProperties } from './OutlookProperties.js';
import { PSTMessage } from './PSTMessage.class.js';
import { PSTFile } from './PSTFile.class.js';
import { RecurrencePattern } from './RecurrencePattern.class.js';
import type { PLNode } from './PLNode.js';
import type { PropertyFinder } from './PAUtil.js';
import type { PLSubNode } from './PLSubNode.js';
import type { RootProvider } from './RootProvider.js';
import type { IPSTTask } from './IPSTTask.js';

export class PSTTask extends PSTMessage implements IPSTTask {
  /**
   * Creates an instance of PSTTask.
   * @internal
   * @param {PSTFile} rootProvider
   * @param {DescriptorIndexNode} descriptorIndexNode
   * @param {Map<number, PSTDescriptorItem>} [localDescriptorItems]
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
   * Specifies the status of the user's progress on the task.
   * https://msdn.microsoft.com/en-us/library/office/cc842120.aspx
   * @readonly
   * @type {number}
   */
  public get taskStatus(): number {
    return this.getIntItem(
      this._rootProvider.getNameToIdMapItem(
        OutlookProperties.PidLidTaskStatus,
        PSTFile.PSETID_Task
      )
    )
  }

  /**
   * Indicates the progress the user has made on a task.
   * https://msdn.microsoft.com/en-us/library/office/cc839932.aspx
   * @readonly
   * @type {number}
   */
  public get percentComplete(): number {
    return this.getDoubleItem(
      this._rootProvider.getNameToIdMapItem(
        OutlookProperties.PidLidPercentComplete,
        PSTFile.PSETID_Task
      )
    )
  }

  /**
   * Specifies the date when the user completes the task.
   * https://msdn.microsoft.com/en-us/library/office/cc815753.aspx
   * @readonly
   * @type {Date}
   */
  public get taskDateCompleted(): Date | null {
    return this.getDateItem(
      this._rootProvider.getNameToIdMapItem(
        OutlookProperties.PidLidTaskDateCompleted,
        PSTFile.PSETID_Task
      )
    )
  }

  /**
   * Indicates the number of minutes that the user performed a task.
   * https://msdn.microsoft.com/en-us/library/office/cc842253.aspx
   * @readonly
   * @type {number}
   */
  public get taskActualEffort(): number {
    return this.getIntItem(
      this._rootProvider.getNameToIdMapItem(
        OutlookProperties.PidLidTaskActualEffort,
        PSTFile.PSETID_Task
      )
    )
  }

  /**
   * Indicates the amount of time, in minutes, that the user expects to perform a task.
   * https://msdn.microsoft.com/en-us/library/office/cc842485.aspx
   * @readonly
   * @type {number}
   */
  public get taskEstimatedEffort(): number {
    return this.getIntItem(
      this._rootProvider.getNameToIdMapItem(
        OutlookProperties.PidLidTaskEstimatedEffort,
        PSTFile.PSETID_Task
      )
    )
  }

  /**
   * Indicates which copy is the latest update of a task.
   * https://msdn.microsoft.com/en-us/library/office/cc815510.aspx
   * @readonly
   * @type {number}
   */
  public get taskVersion(): number {
    return this.getIntItem(
      this._rootProvider.getNameToIdMapItem(
        OutlookProperties.PidLidTaskVersion,
        PSTFile.PSETID_Task
      )
    )
  }

  /**
   * Indicates the task is complete.
   * https://msdn.microsoft.com/en-us/library/office/cc839514.aspx
   * @readonly
   * @type {boolean}
   */
  public get isTaskComplete(): boolean {
    return this.getBooleanItem(
      this._rootProvider.getNameToIdMapItem(
        OutlookProperties.PidLidTaskComplete,
        PSTFile.PSETID_Task
      )
    )
  }

  /**
   * Contains the name of the task owner.
   * https://msdn.microsoft.com/en-us/library/office/cc842363.aspx
   * @readonly
   * @type {string}
   */
  public get taskOwner(): string {
    return this.getStringItem(
      this._rootProvider.getNameToIdMapItem(
        OutlookProperties.PidLidTaskOwner,
        PSTFile.PSETID_Task
      )
    )
  }

  /**
   * Names the user who was last assigned the task.
   * https://msdn.microsoft.com/en-us/library/office/cc815865.aspx
   * @readonly
   * @type {string}
   */
  public get taskAssigner(): string {
    return this.getStringItem(
      this._rootProvider.getNameToIdMapItem(
        OutlookProperties.PidLidTaskAssigner,
        PSTFile.PSETID_Task
      )
    )
  }

  /**
   * Names the most recent user who was the task owner.
   * https://msdn.microsoft.com/en-us/library/office/cc842278.aspx
   * @readonly
   * @type {string}
   */
  public get taskLastUser(): string {
    return this.getStringItem(
      this._rootProvider.getNameToIdMapItem(
        OutlookProperties.PidLidTaskLastUser,
        PSTFile.PSETID_Task
      )
    )
  }

  /**
   * Provides an aid to custom sorting tasks.
   * https://msdn.microsoft.com/en-us/library/office/cc765654.aspx
   * @readonly
   * @type {number}
   */
  public get taskOrdinal(): number {
    return this.getIntItem(
      this._rootProvider.getNameToIdMapItem(
        OutlookProperties.PidLidTaskOrdinal,
        PSTFile.PSETID_Task
      )
    )
  }

  /**
   * Indicates whether the task includes a recurrence pattern.
   * https://msdn.microsoft.com/en-us/library/office/cc765875.aspx
   * @type {boolean}
   */
  public get isTaskRecurring(): boolean {
    return this.getBooleanItem(
      this._rootProvider.getNameToIdMapItem(
        OutlookProperties.PidLidTaskFRecurring,
        PSTFile.PSETID_Task
      )
    )
  }

  /**
   * https://docs.microsoft.com/en-us/office/client-developer/outlook/mapi/pidlidtaskrecurrence-canonical-property
   * @type {RecurrencePattern}
   */
  public get taskRecurrencePattern(): RecurrencePattern | null {
    const recurrenceBLOB = this.getBinaryItem(
      this._rootProvider.getNameToIdMapItem(
        OutlookProperties.PidLidTaskRecurrence,
        PSTFile.PSETID_Task
      )
    )
    return recurrenceBLOB && new RecurrencePattern(recurrenceBLOB)
  }

  /**
   * https://docs.microsoft.com/en-us/office/client-developer/outlook/mapi/pidlidtaskdeadoccurrence-canonical-property
   * @type {boolean}
   */
  public get taskDeadOccurrence(): boolean {
    return this.getBooleanItem(
      this._rootProvider.getNameToIdMapItem(
        OutlookProperties.PidLidTaskDeadOccurrence,
        PSTFile.PSETID_Task
      )
    )
  }

  /**
   * Indicates the role of the current user relative to the task.
   * https://msdn.microsoft.com/en-us/library/office/cc842113.aspx
   * @readonly
   * @type {number}
   */
  public get taskOwnership(): number {
    return this.getIntItem(
      this._rootProvider.getNameToIdMapItem(
        OutlookProperties.PidLidTaskOwnership,
        PSTFile.PSETID_Task
      )
    )
  }

  /**
   * Indicates the acceptance state of the task.
   * https://msdn.microsoft.com/en-us/library/office/cc839689.aspx
   * @readonly
   * @type {number}
   */
  public get acceptanceState(): number {
    return this.getIntItem(
      this._rootProvider.getNameToIdMapItem(
        OutlookProperties.PidLidTaskAcceptanceState,
        PSTFile.PSETID_Task
      )
    )
  }

  /**
   * JSON stringify the object properties.
   * @returns {string}
   */
  public toJSON(): any {
    const clone = Object.assign(
      {
        messageClass: this.messageClass,
        subject: this.subject,
        importance: this.importance,
        transportMessageHeaders: this.transportMessageHeaders,
        taskStatus: this.taskStatus,
        percentComplete: this.percentComplete,
        taskDateCompleted: this.taskDateCompleted,
        taskActualEffort: this.taskActualEffort,
        taskEstimatedEffort: this.taskEstimatedEffort,
        taskVersion: this.taskVersion,
        isTaskComplete: this.isTaskComplete,
        taskOwner: this.taskOwner,
        taskAssigner: this.taskAssigner,
        taskLastUser: this.taskLastUser,
        taskOrdinal: this.taskOrdinal,
        isTaskRecurring: this.isTaskRecurring,
        taskOwnership: this.taskOwnership,
        acceptanceState: this.acceptanceState,
      },
      this
    )
    return clone
  }
}
