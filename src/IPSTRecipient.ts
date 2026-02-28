import type { IPSTObject } from "./IPSTObject.js";

export interface IPSTRecipient extends IPSTObject {
  /**
   * Contains the recipient type for a message recipient.
   * https://msdn.microsoft.com/en-us/library/office/cc839620.aspx
   * 
   * - {@link Consts.MAPI_ORIG} (0)
   * - {@link Consts.MAPI_TO} (1)
   * - {@link Consts.MAPI_CC} (2)
   * - {@link Consts.MAPI_BCC} (3)
   * 
   * @readonly
    */
  get recipientType(): number;

  /**
   * Contains the messaging user's e-mail address type, such as SMTP.
   * https://msdn.microsoft.com/en-us/library/office/cc815548.aspx
   * @readonly
    */
  get addrType(): string;

  /**
   * Contains the messaging user's e-mail address.
   * https://msdn.microsoft.com/en-us/library/office/cc842372.aspx
   * @readonly
    */
  get emailAddress(): string;

  /**
   * Specifies a bit field that describes the recipient status.
   * https://msdn.microsoft.com/en-us/library/office/cc815629.aspx
   * @readonly
    */
  get recipientFlags(): number;

  /**
   * Specifies the location of the current recipient in the recipient table.
   * https://msdn.microsoft.com/en-us/library/ee201359(v=exchg.80).aspx
   * @readonly
    */
  get recipientOrder(): number;

  /**
   * Contains the SMTP address for the address book object.
   * https://msdn.microsoft.com/en-us/library/office/cc842421.aspx
   * @readonly
    */
  get smtpAddress(): string;

  /**
   * JSON stringify the object properties.
    */
  toJSON(): any;
}