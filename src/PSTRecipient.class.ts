/* eslint-disable @typescript-eslint/no-explicit-any */
import { OutlookProperties } from './OutlookProperties'
import { PropertyFinder } from './PAUtil';
import { PLNode } from './PLNode';
import { PLSubNode } from './PLSubNode';
import { PSTObject } from './PSTObject.class'
import { RootProvider } from './RootProvider';

// Class containing recipient information
export class PSTRecipient extends PSTObject {
  /**
   * Creates an instance of PSTRecipient.
   * @internal
   * @param {Map<number, PSTTableItem>} recipientDetails
   * @memberof PSTRecipient
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
   * Contains the recipient type for a message recipient.
   * https://msdn.microsoft.com/en-us/library/office/cc839620.aspx
   * 
   * - {@link Consts.MAPI_ORIG} (0)
   * - {@link Consts.MAPI_TO} (1)
   * - {@link Consts.MAPI_CC} (2)
   * - {@link Consts.MAPI_BCC} (3)
   * 
   * @readonly
   * @type {number}
   * @memberof PSTMessage
   */
  public get recipientType(): number {
    return this.getIntItem(OutlookProperties.PR_RECIPIENT_TYPE)
  }

  /**
   * Contains the messaging user's e-mail address type, such as SMTP.
   * https://msdn.microsoft.com/en-us/library/office/cc815548.aspx
   * @readonly
   * @type {string}
   * @memberof PSTMessage
   */
  public get addrType(): string {
    return this.getStringItem(OutlookProperties.PR_ADDRTYPE)
  }

  /**
   * Contains the messaging user's e-mail address.
   * https://msdn.microsoft.com/en-us/library/office/cc842372.aspx
   * @readonly
   * @type {string}
   * @memberof PSTMessage
   */
  public get emailAddress(): string {
    return this.getStringItem(OutlookProperties.PR_EMAIL_ADDRESS)
  }

  /**
   * Specifies a bit field that describes the recipient status.
   * https://msdn.microsoft.com/en-us/library/office/cc815629.aspx
   * @readonly
   * @type {number}
   * @memberof PSTRecipient
   */
  public get recipientFlags(): number {
    return this.getIntItem(OutlookProperties.PR_RECIPIENT_FLAGS)
  }

  /**
   * Specifies the location of the current recipient in the recipient table.
   * https://msdn.microsoft.com/en-us/library/ee201359(v=exchg.80).aspx
   * @readonly
   * @type {number}
   * @memberof PSTRecipient
   */
  public get recipientOrder(): number {
    return this.getIntItem(OutlookProperties.PidTagRecipientOrder)
  }

  /**
   * Contains the SMTP address for the address book object.
   * https://msdn.microsoft.com/en-us/library/office/cc842421.aspx
   * @readonly
   * @type {string}
   * @memberof PSTRecipient
   */
  public get smtpAddress(): string {
    // If the recipient address type is SMTP, we can simply return the recipient address.
    const addressType: string = this.addrType
    if (addressType != null && addressType.toLowerCase() === 'smtp') {
      const addr: string = this.emailAddress
      if (addr != null && addr.length != 0) {
        return addr
      }
    }
    // Otherwise, we have to hope the SMTP address is present as the PidTagPrimarySmtpAddress property.
    return this.getStringItem(OutlookProperties.PR_SMTP_ADDRESS)
  }

  /**
   * JSON stringify the object properties.
   * @returns {string}
   * @memberof PSTRecipient
   */
  public toJSON(): any {
    const clone = Object.assign(
      {
        displayName: this.displayName,
        smtpAddress: this.smtpAddress,
        recipientType: this.recipientType,
        addrType: this.addrType,
        emailAddress: this.emailAddress,
        recipientFlags: this.recipientFlags,
        recipientOrder: this.recipientOrder,
      },
      this
    )
    return clone
  }
}
