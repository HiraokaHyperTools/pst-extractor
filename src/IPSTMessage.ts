import type { IPSTAttachment } from "./IPSTAttachment.js";
import type { IPSTObject } from "./IPSTObject.js";
import type { IPSTRecipient } from "./IPSTRecipient.js";
import type { PUSubNode } from "./PUSubNode.js";

export interface IPSTMessage extends IPSTObject {
  /**
   * The message is marked as having been read.
   * https://msdn.microsoft.com/en-us/library/ee160304(v=exchg.80).aspx
   * @readonly
   */
  get isRead(): boolean;

  /**
   * The outgoing message has not been modified since the first time that it was saved; the incoming message has not been modified since it was delivered.
   * https://msdn.microsoft.com/en-us/library/ee160304(v=exchg.80).aspx
   * @readonly
   */
  get isUnmodified(): boolean;

  /**
   * The message is marked for sending as a result of a call to the RopSubmitMessage ROP
   * https://msdn.microsoft.com/en-us/library/ee160304(v=exchg.80).aspx
   * @readonly
   */
  get isSubmitted(): boolean;

  /**
   * The message is still being composed. It is saved, but has not been sent.
   * https://msdn.microsoft.com/en-us/library/ee160304(v=exchg.80).aspx
   * @readonly
   */
  get isUnsent(): boolean;

  /**
   * The message has at least one attachment.
   * https://msdn.microsoft.com/en-us/library/ee160304(v=exchg.80).aspx
   * @readonly
   */
  get hasAttachments(): boolean;

  /**
   * The user receiving the message was also the user who sent the message.
   * https://msdn.microsoft.com/en-us/library/ee160304(v=exchg.80).aspx
   * @readonly
   */
  get isFromMe(): boolean;

  /**
   * The message is an FAI message.  An FAI Message object is used to store a variety of settings and
   * auxiliary data, including forms, views, calendar options, favorites, and category lists.
   * https://msdn.microsoft.com/en-us/library/ee160304(v=exchg.80).aspx
   * @readonly
   */
  get isAssociated(): boolean;

  /**
   * The message includes a request for a resend operation with a nondelivery report.
   * https://msdn.microsoft.com/en-us/library/ee160304(v=exchg.80).aspx
   * @readonly
   */
  get isResent(): boolean;

  /**
   * Get the recipients table.
   * @readonly
   */
  getNumberOfRecipients(): Promise<number>;

  /**
   * Get specific recipient.
   */
  getRecipient(recipientNumber: number): Promise<IPSTRecipient>;

  getRecipients(): Promise<IPSTRecipient[]>;

  /**
   * Contains TRUE if a message sender wants notification of non-receipt for a specified recipient.
   * https://msdn.microsoft.com/en-us/library/office/cc979208.aspx
   * @readonly
   */
  get isNonReceiptNotificationRequested(): boolean;

  /**
   * Contains TRUE if a message sender wants notification of non-deliver for a specified recipient.
   * https://msdn.microsoft.com/en-us/library/ms987568(v=exchg.65).aspx
   * @readonly
   */
  get isOriginatorNonDeliveryReportRequested(): boolean;

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
   * Plain text message body.
   * https://msdn.microsoft.com/en-us/library/office/cc765874.aspx
   * @readonly
   */
  get body(): string;

  /**
   * Plain text body prefix.
   * @readonly
   */
  get bodyPrefix(): string;

  /**
   * Contains the Rich Text Format (RTF) version of the message text, usually in compressed form.
   * https://technet.microsoft.com/en-us/library/cc815911
   * @readonly
   */
  get bodyRTF(): string;

  /**
   * Contains the cyclical redundancy check (CRC) computed for the message text.
   * https://technet.microsoft.com/en-us/library/cc815532(v=office.15).aspx
   * @readonly
   */
  get rtfSyncBodyCRC(): number;

  /**
   * Contains a count of the significant characters of the message text.
   * https://msdn.microsoft.com/en-us/library/windows/desktop/cc842324.aspx
   * @readonly
   */
  get rtfSyncBodyCount(): number;

  /**
   * Contains significant characters that appear at the beginning of the message text.
   * https://technet.microsoft.com/en-us/library/cc815400(v=office.15).aspx
   * @readonly
   */
  get rtfSyncBodyTag(): string;

  /**
   * Contains a count of the ignorable characters that appear before the significant characters of the message.
   * https://msdn.microsoft.com/en-us/magazine/cc842437.aspx
   * @readonly
   */
  get rtfSyncPrefixCount(): number;

  /**
   * Contains a count of the ignorable characters that appear after the significant characters of the message.
   * https://msdn.microsoft.com/en-us/magazine/cc765795.aspx
   * @readonly
   */
  get rtfSyncTrailingCount(): number;

  /**
   * Contains the HTML version of the message text.
   * @readonly
   */
  get bodyHTML(): string;

  /**
   * Number of attachments by counting rows in attachment table.
   * @readonly
   */
  getNumberOfAttachments(): Promise<number>;

  /**
   * Get specific attachment from table using index.
   */
  getAttachment(attachmentNumber: number): Promise<IPSTAttachment>;

  getAttachments(): Promise<IPSTAttachment[]>;

  /**
   * Importance of email (sender determined)
   * https://msdn.microsoft.com/en-us/library/cc815346(v=office.12).aspx
   * @readonly
   */
  get importance(): number;

  /**
   * Contains a text string that identifies the sender-defined message class, such as IPM.Note.
   * https://msdn.microsoft.com/en-us/library/office/cc765765.aspx
   * @readonly
   */
  get messageClass(): string;

  /**
   * Contains the full subject of a message.
   * https://technet.microsoft.com/en-us/library/cc815720
   * @readonly
   */
  get subject(): string;

  /**
   * Contains the date and time the message sender submitted a message.
   * https://technet.microsoft.com/en-us/library/cc839781
   * @readonly
   */
  get clientSubmitTime(): Date | null;

  /**
   * Contains the display name of the messaging user who receives the message.
   * https://msdn.microsoft.com/en-us/library/office/cc840015.aspx
   * @readonly
   */
  get receivedByName(): string;

  /**
   * Contains the display name for the messaging user represented by the sender.
   * https://msdn.microsoft.com/en-us/library/office/cc842405.aspx
   * @readonly
   */
  get sentRepresentingName(): string;

  /**
   * Contains the address type for the messaging user who is represented by the sender.
   * https://msdn.microsoft.com/en-us/library/office/cc839677.aspx
   * @readonly
   */
  get sentRepresentingAddressType(): string;

  /**
   * Contains the e-mail address for the messaging user who is represented by the sender.
   * https://msdn.microsoft.com/en-us/library/office/cc839552.aspx
   * @readonly
   */
  get sentRepresentingEmailAddress(): string;

  /**
   * Contains the topic of the first message in a conversation thread.
   * https://technet.microsoft.com/en-us/windows/cc839841
   * @readonly
   */
  get conversationTopic(): string;

  /**
   * Contains the e-mail address type, such as SMTP, for the messaging user who actually receives the message.
   * https://technet.microsoft.com/en-us/library/cc765641(v=office.14)
   * @readonly
   */
  get receivedByAddressType(): string;

  /**
   * Contains the e-mail address for the messaging user who receives the message.
   * https://technet.microsoft.com/en-us/library/cc839550(v=office.14)
   * @readonly
   */
  get receivedByAddress(): string;

  /**
   * Contains transport-specific message envelope information.
   * https://technet.microsoft.com/en-us/library/cc815628
   * @readonly
   */
  get transportMessageHeaders(): string;

  /**
   * Acknowledgment mode Integer 32-bit signed
   */
  get acknowledgementMode(): number;

  /**
   * Contains TRUE if a message sender requests a delivery report for a particular recipient from the messaging system before the message is placed in the message store.
   * https://msdn.microsoft.com/en-us/library/office/cc765845.aspx
   * @readonly
   */
  get originatorDeliveryReportRequested(): boolean;

  /**
   * Contains the relative priority of a message.
   * https://msdn.microsoft.com/en-us/library/office/cc765646.aspx
   * @readonly
   */
  get priority(): number;

  /**
   * Contains TRUE if a message sender wants the messaging system to generate a read report when the recipient has read a message.
   * https://msdn.microsoft.com/en-us/library/office/cc842094.aspx
   * @readonly
   */
  get readReceiptRequested(): boolean;

  /**
   * Specifies whether adding additional recipients, when forwarding the message, is prohibited for the e-mail message.
   * https://msdn.microsoft.com/en-us/library/office/cc979216.aspx
   * @readonly
   */
  get recipientReassignmentProhibited(): boolean;

  /**
   * Contains the sensitivity value assigned by the sender of the first version of a message that is, the message before being forwarded or replied to.
   * https://msdn.microsoft.com/en-us/library/cc839694(office.12).aspx
   * @readonly
   */
  get originalSensitivity(): number;

  /**
   * Contains a value that indicates the message sender's opinion of the sensitivity of a message.
   * https://msdn.microsoft.com/en-us/library/office/cc839518.aspx
   * @readonly
   */
  get sensitivity(): number;

  /**
   * Contains the search key for the messaging user represented by the sender.
   * https://msdn.microsoft.com/en-us/magazine/cc842068.aspx
   * @readonly
   */
  get pidTagSentRepresentingSearchKey(): Uint8Array | null;

  /**
   * Contains the display name for the messaging user who is represented by the receiving user.
   * https://technet.microsoft.com/en-us/library/cc842260.aspx
   * @readonly
   */
  get rcvdRepresentingName(): string;

  /**
   * Contains the subject of an original message for use in a report about the message.
   * https://msdn.microsoft.com/en-us/library/office/cc842182.aspx
   * @readonly
   */
  get originalSubject(): string;

  /**
   * Contains a list of display names for recipients that are to get a reply.
   * https://msdn.microsoft.com/en-us/library/windows/desktop/cc815850.aspx
   * @readonly
   */
  get replyRecipientNames(): string;

  /**
   * Contains TRUE if this messaging user is specifically named as a primary (To) recipient of this message and is not part of a distribution list.
   * https://technet.microsoft.com/en-us/library/cc815755
   * @readonly
   */
  get messageToMe(): boolean;

  /**
   * Contains TRUE if this messaging user is specifically named as a carbon copy (CC) recipient of this message and is not part of a distribution list.
   * https://msdn.microsoft.com/en-us/library/office/cc839713.aspx
   * @readonly
   */
  get messageCcMe(): boolean;

  /**
   * Contains TRUE if this messaging user is specifically named as a primary (To), carbon copy (CC), or blind carbon copy (BCC) recipient of this message and is not part of a distribution list.
   * https://msdn.microsoft.com/en-us/library/office/cc842268.aspx
   * @readonly
   */
  get messageRecipMe(): boolean;

  /**
   * Contains TRUE if the message sender wants a response to a meeting request.
   * https://msdn.microsoft.com/en-us/library/office/cc839921.aspx
   * @readonly
   */
  get responseRequested(): boolean;

  /**
   * Contains the display names of any carbon copy (CC) recipients of the original message.
   * https://msdn.microsoft.com/en-us/magazine/cc815841(v=office.14).aspx
   * @readonly
   */
  get originalDisplayBcc(): string;

  /**
   * Contains the display names of any carbon copy (CC) recipients of the original message.
   * https://msdn.microsoft.com/en-us/magazine/cc815841(v=office.14).aspx
   * @readonly
   */
  get originalDisplayCc(): string;

  /**
   * Contains the display names of the primary (To) recipients of the original message.
   * https://msdn.microsoft.com/en-us/magazine/cc842235(v=office.14).aspx
   * @readonly
   */
  get originalDisplayTo(): string;

  /**
   * Contains the address type for the messaging user who is represented by the user actually receiving the message.
   * https://msdn.microsoft.com/en-us/library/office/cc842447.aspx
   * @readonly
   */
  get rcvdRepresentingAddrtype(): string;

  /**
   * Contains the e-mail address for the messaging user who is represented by the receiving user.
   * https://msdn.microsoft.com/en-us/library/office/cc815875.aspx
   * @readonly
   */
  get rcvdRepresentingEmailAddress(): string;

  /**
   * Contains TRUE if a message sender requests a reply from a recipient.
   * https://msdn.microsoft.com/en-us/library/office/cc815286.aspx
   * @readonly
   */
  get isReplyRequested(): boolean;

  /**
   * Contains the message sender's entry identifier.
   * https://msdn.microsoft.com/en-us/library/office/cc815625.aspx
   * @readonly
   */
  get senderEntryId(): Uint8Array | null;

  /**
   * Contains the message sender's display name.
   * https://msdn.microsoft.com/en-us/library/office/cc815457.aspx
   * @readonly
   */
  get senderName(): string;

  /**
   * Contains the message sender's e-mail address type.
   * https://msdn.microsoft.com/en-us/library/office/cc815748.aspx
   * @readonly
   */
  get senderAddrtype(): string;

  /**
   * Contains the message sender's e-mail address.
   * https://msdn.microsoft.com/en-us/library/office/cc839670.aspx
   * @readonly
   */
  get senderEmailAddress(): string;

  /**
   * Contains the sum, in bytes, of the sizes of all properties on a message object
   * https://technet.microsoft.com/en-us/library/cc842471
   * @readonly
   */
  get messageSize(): Long;

  /**
   * A number associated with an item in a message store.
   * https://msdn.microsoft.com/en-us/library/office/cc815718.aspx
   * @readonly
   */
  get internetArticleNumber(): number;

  /**
   * Contains a string that names the first server that is used to send the message.
   * https://msdn.microsoft.com/en-us/library/office/cc815413.aspx
   * @readonly
   */
  get primarySendAccount(): string;

  /**
   * Specifies the server that a client is currently attempting to use to send e-mail.
   * https://technet.microsoft.com/en-us/library/cc842327(v=office.14)
   * @readonly
   */
  get nextSendAcct(): string;

  /**
   * Contains the type of an object.
   * https://msdn.microsoft.com/en-us/library/office/cc815487.aspx
   * @readonly
   */
  get objectType(): number;

  /**
   * Contains TRUE if a client application wants MAPI to delete the associated message after submission.
   * https://msdn.microsoft.com/en-us/library/office/cc842353.aspx
   * @readonly
   */
  get deleteAfterSubmit(): boolean;

  /**
   * Contains TRUE if some transport provider has already accepted responsibility for delivering the message to this recipient, and FALSE if the MAPI spooler considers that this transport provider should accept responsibility.
   * https://msdn.microsoft.com/en-us/library/office/cc765767.aspx
   * @readonly
   */
  get responsibility(): boolean;

  /**
   * Contains TRUE if the PR_RTF_COMPRESSED (PidTagRtfCompressed) property has the same text content as the PR_BODY (PidTagBody) property for this message.
   * https://msdn.microsoft.com/en-us/library/office/cc765844.aspx
   * @readonly
   */
  get isRTFInSync(): boolean;

  /**
   * Contains an ASCII list of the display names of any blind carbon copy (BCC) message recipients, separated by semicolons (;).
   * https://msdn.microsoft.com/en-us/library/office/cc815730.aspx
   * @readonly
   */
  get displayBCC(): string;

  /**
   * Contains an ASCII list of the display names of any carbon copy (CC) message recipients, separated by semicolons (;).
   * https://msdn.microsoft.com/en-us/library/office/cc765528.aspx
   * @readonly
   */
  get displayCC(): string;

  /**
   * Contains a list of the display names of the primary (To) message recipients, separated by semicolons (;).
   * https://msdn.microsoft.com/en-us/library/office/cc839687.aspx
   * @readonly
   */
  get displayTo(): string;

  /**
   * Contains the date and time when a message was delivered.
   * https://msdn.microsoft.com/en-us/library/office/cc841961.aspx
   * @readonly
   */
  get messageDeliveryTime(): Date | null;

  /**
   * Corresponds to the message ID field as specified in [RFC2822].
   * https://msdn.microsoft.com/en-us/library/office/cc839521.aspx
   * @readonly
   */
  get internetMessageId(): string;

  /**
   * Contains the original message's PR_INTERNET_MESSAGE_ID (PidTagInternetMessageId) property value.
   * https://msdn.microsoft.com/en-us/library/office/cc839776.aspx
   * @readonly
   */
  get inReplyToId(): string;

  /**
   * Contains the value of a Multipurpose Internet Mail Extensions (MIME) message's Return-Path header field. The e-mail address of the message's sender.
   * https://msdn.microsoft.com/en-us/library/office/cc765856.aspx
   * @readonly
   */
  get returnPath(): string;

  /**
   * Contains a number that indicates which icon to use when you display a group of e-mail objects.
   * https://msdn.microsoft.com/en-us/library/office/cc815472.aspx
   * @readonly
   */
  get iconIndex(): number;

  /**
   * Contains the last verb executed.
   * Todo: Helper methods for each flag.
   * https://msdn.microsoft.com/en-us/library/office/cc841968.aspx
   * @readonly
   */
  get lastVerbExecuted(): number;

  /**
   * Contains the time when the last verb was executed.
   * https://msdn.microsoft.com/en-us/library/office/cc839918.aspx
   * @readonly
   */
  get lastVerbExecutionTime(): Date | null;

  /**
   * The URL component name for a message.
   * https://msdn.microsoft.com/en-us/library/office/cc815653.aspx
   * @readonly
   */
  get urlCompName(): string;

  /**
   * Specifies the hide or show status of a folder.
   * https://msdn.microsoft.com/en-us/library/ee159038(v=exchg.80).aspx
   * @readonly
   */
  get attrHidden(): boolean;

  /**
   * Specifies the date on which the user expects work on the task to begin.
   * https://technet.microsoft.com/en-us/library/cc815922(v=office.12).aspx
   * @readonly
   */
  get taskStartDate(): Date | null;

  /**
   * Represents the date when the user expects to complete the task.
   * https://technet.microsoft.com/en-us/library/cc839641(v=office.12).aspx
   * @readonly
   */
  get taskDueDate(): Date | null;

  /**
   * Specifies whether a reminder is set on the object.
   * https://msdn.microsoft.com/en-us/library/office/cc765589.aspx
   * @readonly
   */
  get reminderSet(): boolean;

  /**
   * Specifies the interval, in minutes, between the time when the reminder first becomes overdue and the start time of the calendar object.
   * https://msdn.microsoft.com/en-us/library/office/cc765535.aspx
   * @readonly
   */
  get reminderDelta(): number;

  /**
   * Color categories
   * @readonly
   */
  get colorCategories(): string[];

  /**
   * Contains a computed value derived from other conversation-related properties.
   * https://msdn.microsoft.com/en-us/library/ee204279(v=exchg.80).aspx
   * @readonly
   */
  get conversationId(): Uint8Array | null;

  /**
   * Indicates whether the GUID portion of the PidTagConversationIndex property (section 2.641) is to be used to compute the PidTagConversationId property (section 2.640).
   * https://msdn.microsoft.com/en-us/library/ee218393(v=exchg.80).aspx
   * @readonly
   */
  get isConversationIndexTracking(): boolean;

  /**
   * Contains the messaging user's e-mail address.
   * https://msdn.microsoft.com/en-us/library/office/cc842372.aspx
   * @readonly
   */
  get emailAddress(): string;

  /**
   * Contains the messaging user's e-mail address type, such as SMTP.
   * https://msdn.microsoft.com/en-us/library/office/cc815548.aspx
   * @readonly
   */
  get addrType(): string;

  /**
   * Contains a comment about the purpose or content of an object.
   * https://msdn.microsoft.com/en-us/library/office/cc842022.aspx
   * @readonly
   */
  get comment(): string;

  /**
   * Contains the creation date and time of a message.
   * https://msdn.microsoft.com/en-us/library/office/cc765677.aspx
   * @readonly
   */
  get creationTime(): Date | null;

  /**
   * Contains the date and time when the object or subobject was last modified.
   * https://msdn.microsoft.com/en-us/library/office/cc815689.aspx
   * @readonly
   */
  get modificationTime(): Date | null;

  /**
   * Requests access to the user sub-node of the internal PST structure.
   */
  requestAccessToUserSubNode(): Promise<PUSubNode | undefined>;

  /**
   * JSON stringify the object properties.  Large fields (like body) aren't included.
   */
  toJSON(): any;
}
