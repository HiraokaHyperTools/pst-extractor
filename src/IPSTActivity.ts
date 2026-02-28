import type { IPSTMessage } from "./IPSTMessage.js";

export interface IPSTActivity extends IPSTMessage {
  /**
   * Contains the display name of the journaling application (for example, "MSWord"), and is typically a free-form attribute of a journal message, usually a string.
   * https://msdn.microsoft.com/en-us/library/office/cc839662.aspx
   * @readonly
   */
  get logType(): string;

  /**
   * Represents the start date and time for the journal message.
   * https://msdn.microsoft.com/en-us/library/office/cc842339.aspx
   * @readonly
   */
  get logStart(): Date | null;

  /**
   * Represents the duration, in minutes, of a journal message.
   * https://msdn.microsoft.com/en-us/library/office/cc765536.aspx
   * @readonly
   */
  get logDuration(): number;

  /**
   * Represents the end date and time for the journal message.
   * https://msdn.microsoft.com/en-us/library/office/cc839572.aspx
   * @readonly
   */
  get logEnd(): Date | null;

  /**
   * Contains metadata about the journal.
   * https://msdn.microsoft.com/en-us/library/office/cc815433.aspx
   * @readonly
   */
  get logFlags(): number;

  /**
   * Indicates whether the document was printed during journaling.
   * https://msdn.microsoft.com/en-us/library/office/cc839873.aspx
   * @readonly
   */
  get isDocumentPrinted(): boolean;

  /**
   * Indicates whether the document was saved during journaling.
   * https://msdn.microsoft.com/en-us/library/office/cc815488.aspx
   * @readonly
   */
  get isDocumentSaved(): boolean;

  /**
   * Indicates whether the document was sent to a routing recipient during journaling.
   * https://msdn.microsoft.com/en-us/library/office/cc839558.aspx
   * @readonly
   */
  get isDocumentRouted(): boolean;

  /**
   * Indicates whether the document was sent by e-mail or posted to a server folder during journaling.
   * https://msdn.microsoft.com/en-us/library/office/cc815353.aspx
   * @readonly
   */
  get isDocumentPosted(): boolean;

  /**
   * Describes the activity that is being recorded.
   * https://msdn.microsoft.com/en-us/library/office/cc815500.aspx
   * @readonly
   */
  get logTypeDesc(): string;

  /**
   * JSON stringify the object properties.
   */
  toJSON(): any;
}