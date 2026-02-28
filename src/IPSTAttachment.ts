import type { IPSTMessage } from "./IPSTMessage.js";

export interface IPSTAttachment {
  /**
   * The PR_ATTACH_SIZE property contains the sum, in bytes, of the sizes of all properties on an attachment.
   * https://msdn.microsoft.com/en-us/library/gg156074(v=winembedded.70).aspx
   * @readonly
   */
  get size(): number;

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
   * Get an embedded message.
   * @readonly
   */
  getEmbeddedPSTMessage(): Promise<IPSTMessage | null>;

  /**
   * Get attachment content as binary data
   */
  get fileData(): ArrayBuffer | undefined;

  /**
   * Size of the attachment file itself.
   * https://msdn.microsoft.com/en-us/library/gg154634(v=winembedded.70).aspx
   * @readonly
   */
  get filesize(): number;

  /**
   * Contains an attachment's base file name and extension, excluding path.
   * https://msdn.microsoft.com/en-us/library/office/cc842517.aspx
   * @readonly
   */
  get filename(): string;

  /**
   * Contains a MAPI-defined constant representing the way the contents of an attachment can be accessed.
   * https://msdn.microsoft.com/en-us/library/office/cc815439.aspx
   * 
   * - {@link PSTAttachment.ATTACHMENT_METHOD_BY_VALUE} (1)
   * - {@link PSTAttachment.ATTACHMENT_METHOD_BY_REFERENCE} (2)
   * - {@link PSTAttachment.ATTACHMENT_METHOD_BY_REFERENCE_RESOLVE} (3)
   * - {@link PSTAttachment.ATTACHMENT_METHOD_BY_REFERENCE_ONLY} (4)
   * - {@link PSTAttachment.ATTACHMENT_METHOD_EMBEDDED} (5)
   * - {@link PSTAttachment.ATTACHMENT_METHOD_OLE} (6)
   *
   * @readonly
   */
  get attachMethod(): number;

  /**
   * Contains a number that uniquely identifies the attachment within its parent message.
   * https://msdn.microsoft.com/en-us/library/office/cc841969.aspx
   * @readonly
   */
  get attachNum(): number;

  /**
   * Contains an attachment's long filename and extension, excluding path.
   * https://msdn.microsoft.com/en-us/library/office/cc842157.aspx
   * @readonly
   */
  get longFilename(): string;

  /**
   * Contains an attachment's fully-qualified path and filename.
   * https://msdn.microsoft.com/en-us/library/office/cc839889.aspx
   * @readonly
   */
  get pathname(): string;

  /**
   * Contains an offset, in characters, to use in rendering an attachment within the main message text.
   * https://msdn.microsoft.com/en-us/library/office/cc842381.aspx
   * @readonly
   */
  get renderingPosition(): number;

  /**
   * Contains an attachment's fully-qualified long path and filename.
   * https://msdn.microsoft.com/en-us/library/office/cc815443.aspx
   * @readonly
   */
  get longPathname(): string;

  /**
   * Contains formatting information about a Multipurpose Internet Mail Extensions (MIME) attachment.
   * https://msdn.microsoft.com/en-us/library/office/cc842516.aspx
   * @readonly
   */
  get mimeTag(): string;

  /**
   * Contains the MIME sequence number of a MIME message attachment.
   * https://msdn.microsoft.com/en-us/library/office/cc963256.aspx
   * @readonly
   */
  get mimeSequence(): number;

  /**
   * Contains the content identification header of a Multipurpose Internet Mail Extensions (MIME) message attachment.
   * https://msdn.microsoft.com/en-us/library/office/cc765868.aspx
   * @readonly
   */
  get contentId(): string;

  /**
   * Indicates that this attachment is not available to HTML rendering applications and should be ignored in Multipurpose Internet Mail Extensions (MIME) processing.
   * https://msdn.microsoft.com/en-us/library/office/cc765876.aspx
   * @readonly
   */
  get isAttachmentInvisibleInHtml(): boolean;

  /**
   * Indicates that this attachment is not available to applications rendering in Rich Text Format (RTF) and should be ignored by MAPI.
   * https://msdn.microsoft.com/en-us/library/office/cc765876.aspx
   * @readonly
   */
  get isAttachmentInvisibleInRTF(): boolean;

  /**
   * JSON stringify the object properties.
   */
  toJSON(): any;
}