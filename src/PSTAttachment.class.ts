/* eslint-disable @typescript-eslint/no-explicit-any */
import type { IPSTAttachment } from './IPSTAttachment.js';
import { OutlookProperties } from './OutlookProperties.js';
import type { PropertyFinder } from './PAUtil.js';
import type { PLNode } from './PLNode.js';
import type { PLSubNode } from './PLSubNode.js';
import { PropertyTypeObject } from './PropertyTypeObject.js';
import { PSTFile } from './PSTFile.class.js';
import { PSTMessage } from './PSTMessage.class.js';
import { PSTObject } from './PSTObject.class.js';
import { PSTUtil } from './PSTUtil.class.js';
import type { RootProvider } from './RootProvider.js';

// Class containing attachment information.
export class PSTAttachment extends PSTObject implements IPSTAttachment {
  public static ATTACHMENT_METHOD_NONE = 0
  public static ATTACHMENT_METHOD_BY_VALUE = 1
  public static ATTACHMENT_METHOD_BY_REFERENCE = 2
  public static ATTACHMENT_METHOD_BY_REFERENCE_RESOLVE = 3
  public static ATTACHMENT_METHOD_BY_REFERENCE_ONLY = 4
  public static ATTACHMENT_METHOD_EMBEDDED = 5
  public static ATTACHMENT_METHOD_OLE = 6

  /**
   * Creates an instance of PSTAttachment.
   * @internal
   * @param {PSTFile} rootProvider
   * @param {Map<number, PSTDescriptorItem>} localDescriptorItems
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
   * The PR_ATTACH_SIZE property contains the sum, in bytes, of the sizes of all properties on an attachment.
   * https://msdn.microsoft.com/en-us/library/gg156074(v=winembedded.70).aspx
   * @readonly
   * @type {number}
   */
  public get size(): number {
    return this.getIntItem(OutlookProperties.PR_ATTACH_SIZE)
  }

  /**
   * Contains the creation date and time of a message.
   * https://msdn.microsoft.com/en-us/library/office/cc765677.aspx
   * @readonly
   * @type {Date}
   */
  public get creationTime(): Date | null {
    return this.getDateItem(OutlookProperties.PR_CREATION_TIME)
  }

  /**
   * Contains the date and time when the object or subobject was last modified.
   * https://msdn.microsoft.com/en-us/library/office/cc815689.aspx
   * @readonly
   * @type {Date}
   */
  public get modificationTime(): Date | null {
    return this.getDateItem(OutlookProperties.PR_LAST_MODIFICATION_TIME)
  }

  /**
   * Get an embedded message.
   * @readonly
   * @type {PSTMessage}
   */
  public async getEmbeddedPSTMessage(): Promise<PSTMessage | null> {
    const attachMethod = this._propertyFinder.findByKey(0x3705)?.value;

    try {
      if (true
        && typeof attachMethod === 'number'
        && attachMethod == PSTAttachment.ATTACHMENT_METHOD_EMBEDDED
      ) {
        const attachDataBinary = this._propertyFinder.findByKey(0x3701)?.value;
        if (false) { }
        else if (attachDataBinary instanceof ArrayBuffer) {
          // PT_BINARY

          throw new Error("Currently getEmbeddedPSTMessage and ATTACHMENT_METHOD_EMBEDDED need attachDataBinary to be PT_OBJECT");
        }
        else if (attachDataBinary instanceof PropertyTypeObject) {
          const { subNodeId } = attachDataBinary;

          const subNode = await this._subNode.getChildBy(subNodeId);

          if (subNode === undefined) {
            throw new Error(
              `childNodeId=0x${subNodeId.toString(16)}`
              + ` of ${this._subNode} not found`
            );
          }

          return await this._rootProvider.getItemOf(
            this._node,
            subNode
          );
        }
      }
    } catch (err) {
      console.error(
        'PSTAttachment::embeddedPSTMessage createAppropriatePSTMessageObject failed\n' +
        err
      )
      throw err
    }

    return null
  }

  /**
   * Get attachment content as binary data
   */
  public get fileData(): ArrayBuffer | undefined {
    const attachmentDataObject = this._propertyFinder.findByKey(
      OutlookProperties.PR_ATTACH_DATA_BIN
    );
    if (attachmentDataObject !== undefined) {
      const { value } = attachmentDataObject;
      if (value instanceof ArrayBuffer) {
        return value;
      }
    }
    return undefined;
  }

  /**
   * Size of the attachment file itself.
   * https://msdn.microsoft.com/en-us/library/gg154634(v=winembedded.70).aspx
   * @readonly
   * @type {number}
   */
  public get filesize(): number {
    const attachmentDataObject = this._propertyFinder.findByKey(
      OutlookProperties.PR_ATTACH_DATA_BIN
    );
    if (attachmentDataObject !== undefined) {
      const { value } = attachmentDataObject;
      if (value instanceof ArrayBuffer) {
        return value.byteLength;
      }
      else if (value instanceof PropertyTypeObject) {
        return value.size;
      }
    }
    return 0
  }

  /**
   * Contains an attachment's base file name and extension, excluding path.
   * https://msdn.microsoft.com/en-us/library/office/cc842517.aspx
   * @readonly
   * @type {string}
   */
  public get filename(): string {
    return this.getStringItem(OutlookProperties.PR_ATTACH_FILENAME)
  }

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
   * @type {number}
   */
  public get attachMethod(): number {
    return this.getIntItem(OutlookProperties.PR_ATTACH_METHOD)
  }

  /**
   * Contains a number that uniquely identifies the attachment within its parent message.
   * https://msdn.microsoft.com/en-us/library/office/cc841969.aspx
   * @readonly
   * @type {number}
   */
  public get attachNum(): number {
    return this.getIntItem(OutlookProperties.PR_ATTACH_NUM)
  }

  /**
   * Contains an attachment's long filename and extension, excluding path.
   * https://msdn.microsoft.com/en-us/library/office/cc842157.aspx
   * @readonly
   * @type {string}
   */
  public get longFilename(): string {
    return this.getStringItem(OutlookProperties.PR_ATTACH_LONG_FILENAME)
  }

  /**
   * Contains an attachment's fully-qualified path and filename.
   * https://msdn.microsoft.com/en-us/library/office/cc839889.aspx
   * @readonly
   * @type {string}
   */
  public get pathname(): string {
    return this.getStringItem(OutlookProperties.PR_ATTACH_PATHNAME)
  }

  /**
   * Contains an offset, in characters, to use in rendering an attachment within the main message text.
   * https://msdn.microsoft.com/en-us/library/office/cc842381.aspx
   * @readonly
   * @type {number}
   */
  public get renderingPosition(): number {
    return this.getIntItem(OutlookProperties.PR_RENDERING_POSITION)
  }

  /**
   * Contains an attachment's fully-qualified long path and filename.
   * https://msdn.microsoft.com/en-us/library/office/cc815443.aspx
   * @readonly
   * @type {string}
   */
  public get longPathname(): string {
    return this.getStringItem(OutlookProperties.PR_ATTACH_LONG_PATHNAME)
  }

  /**
   * Contains formatting information about a Multipurpose Internet Mail Extensions (MIME) attachment.
   * https://msdn.microsoft.com/en-us/library/office/cc842516.aspx
   * @readonly
   * @type {string}
   */
  public get mimeTag(): string {
    return this.getStringItem(OutlookProperties.PR_ATTACH_MIME_TAG)
  }

  /**
   * Contains the MIME sequence number of a MIME message attachment.
   * https://msdn.microsoft.com/en-us/library/office/cc963256.aspx
   * @readonly
   * @type {number}
   */
  public get mimeSequence(): number {
    return this.getIntItem(OutlookProperties.PR_ATTACH_MIME_SEQUENCE)
  }

  /**
   * Contains the content identification header of a Multipurpose Internet Mail Extensions (MIME) message attachment.
   * https://msdn.microsoft.com/en-us/library/office/cc765868.aspx
   * @readonly
   * @type {string}
   */
  public get contentId(): string {
    return this.getStringItem(OutlookProperties.PR_ATTACH_CONTENT_ID)
  }

  /**
   * Indicates that this attachment is not available to HTML rendering applications and should be ignored in Multipurpose Internet Mail Extensions (MIME) processing.
   * https://msdn.microsoft.com/en-us/library/office/cc765876.aspx
   * @readonly
   * @type {boolean}
   */
  public get isAttachmentInvisibleInHtml(): boolean {
    const actionFlag = this.getIntItem(OutlookProperties.PR_ATTACH_FLAGS)
    return (actionFlag & 0x1) > 0
  }

  /**
   * Indicates that this attachment is not available to applications rendering in Rich Text Format (RTF) and should be ignored by MAPI.
   * https://msdn.microsoft.com/en-us/library/office/cc765876.aspx
   * @readonly
   * @type {boolean}
   */
  public get isAttachmentInvisibleInRTF(): boolean {
    const actionFlag = this.getIntItem(OutlookProperties.PR_ATTACH_FLAGS)
    return (actionFlag & 0x2) > 0
  }

  /**
   * JSON stringify the object properties.
   * @returns {string}
   */
  public toJSON(): any {
    const clone = Object.assign(
      {
        size: this.size,
        creationTime: this.creationTime,
        modificationTime: this.modificationTime,
        filename: this.filename,
        attachMethod: this.attachMethod,
        attachNum: this.attachNum,
        longFilename: this.longFilename,
        pathname: this.pathname,
        renderingPosition: this.renderingPosition,
        longPathname: this.longPathname,
        mimeTag: this.mimeTag,
        mimeSequence: this.mimeSequence,
        contentId: this.contentId,
        isAttachmentInvisibleInHtml: this.isAttachmentInvisibleInHtml,
        isAttachmentInvisibleInRTF: this.isAttachmentInvisibleInRTF,
        primaryNodeId: this.primaryNodeId,
      },
      this
    )
    return clone
  }
}
