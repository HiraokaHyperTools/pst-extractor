import type { IRecurrencePattern } from "./IRecurrencePattern.js";

export interface IPSTTask {
  /**
   * Specifies the status of the user's progress on the task.
   * https://msdn.microsoft.com/en-us/library/office/cc842120.aspx
   * @readonly
     */
  get taskStatus(): number;

  /**
   * Indicates the progress the user has made on a task.
   * https://msdn.microsoft.com/en-us/library/office/cc839932.aspx
   * @readonly
     */
  get percentComplete(): number;

  /**
   * Specifies the date when the user completes the task.
   * https://msdn.microsoft.com/en-us/library/office/cc815753.aspx
   * @readonly
     */
  get taskDateCompleted(): Date | null;

  /**
   * Indicates the number of minutes that the user performed a task.
   * https://msdn.microsoft.com/en-us/library/office/cc842253.aspx
   * @readonly
     */
  get taskActualEffort(): number;

  /**
   * Indicates the amount of time, in minutes, that the user expects to perform a task.
   * https://msdn.microsoft.com/en-us/library/office/cc842485.aspx
   * @readonly
     */
  get taskEstimatedEffort(): number;

  /**
   * Indicates which copy is the latest update of a task.
   * https://msdn.microsoft.com/en-us/library/office/cc815510.aspx
   * @readonly
     */
  get taskVersion(): number;

  /**
   * Indicates the task is complete.
   * https://msdn.microsoft.com/en-us/library/office/cc839514.aspx
   * @readonly
     */
  get isTaskComplete(): boolean;

  /**
   * Contains the name of the task owner.
   * https://msdn.microsoft.com/en-us/library/office/cc842363.aspx
   * @readonly
     */
  get taskOwner(): string;

  /**
   * Names the user who was last assigned the task.
   * https://msdn.microsoft.com/en-us/library/office/cc815865.aspx
   * @readonly
     */
  get taskAssigner(): string;

  /**
   * Names the most recent user who was the task owner.
   * https://msdn.microsoft.com/en-us/library/office/cc842278.aspx
   * @readonly
     */
  get taskLastUser(): string;

  /**
   * Provides an aid to custom sorting tasks.
   * https://msdn.microsoft.com/en-us/library/office/cc765654.aspx
   * @readonly
     */
  get taskOrdinal(): number;

  /**
   * Indicates whether the task includes a recurrence pattern.
   * https://msdn.microsoft.com/en-us/library/office/cc765875.aspx
     */
  get isTaskRecurring(): boolean;

  /**
   * https://docs.microsoft.com/en-us/office/client-developer/outlook/mapi/pidlidtaskrecurrence-canonical-property
     */
  get taskRecurrencePattern(): IRecurrencePattern | null;

  /**
   * https://docs.microsoft.com/en-us/office/client-developer/outlook/mapi/pidlidtaskdeadoccurrence-canonical-property
     */
  get taskDeadOccurrence(): boolean;

  /**
   * Indicates the role of the current user relative to the task.
   * https://msdn.microsoft.com/en-us/library/office/cc842113.aspx
   * @readonly
     */
  get taskOwnership(): number;

  /**
   * Indicates the acceptance state of the task.
   * https://msdn.microsoft.com/en-us/library/office/cc839689.aspx
   * @readonly
     */
  get acceptanceState(): number;

  /**
   * JSON stringify the object properties.
    */
  toJSON(): any;
}