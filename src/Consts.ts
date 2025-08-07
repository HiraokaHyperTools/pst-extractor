
// from `C:\Program Files (x86)\Windows Kits\10\Include\10.0.20348.0\um\WabDefs.h`

// #define NO_ATTACHMENT           ((ULONG) 0x00000000)
// #define ATTACH_BY_VALUE         ((ULONG) 0x00000001)
// #define ATTACH_BY_REFERENCE     ((ULONG) 0x00000002)
// #define ATTACH_BY_REF_RESOLVE   ((ULONG) 0x00000003)
// #define ATTACH_BY_REF_ONLY      ((ULONG) 0x00000004)
// #define ATTACH_EMBEDDED_MSG     ((ULONG) 0x00000005)
// #define ATTACH_OLE              ((ULONG) 0x00000006)

// #define MAPI_ORIG   0           /* Recipient is message originator          */
// #define MAPI_TO     1           /* Recipient is a primary recipient         */
// #define MAPI_CC     2           /* Recipient is a copy recipient            */
// #define MAPI_BCC    3           /* Recipient is blind copy recipient        */
// #define MAPI_P1     0x10000000  /* Recipient is a P1 resend recipient       */
// #define MAPI_SUBMITTED 0x80000000 /* Recipient is already processed         */

export class Consts {
  public static readonly NO_ATTACHMENT = 0x00000000 as const;
  public static readonly ATTACH_BY_VALUE = 0x00000001 as const;
  public static readonly ATTACH_BY_REFERENCE = 0x00000002 as const;
  public static readonly ATTACH_BY_REF_RESOLVE = 0x00000003 as const;
  public static readonly ATTACH_BY_REF_ONLY = 0x00000004 as const;
  public static readonly ATTACH_EMBEDDED_MSG = 0x00000005 as const;
  public static readonly ATTACH_OLE = 0x00000006 as const;

  /**
   * Recipient is message originator
   */
  public static readonly MAPI_ORIG = 0 as const;

  /**
   * Recipient is a primary recipient
   */
  public static readonly MAPI_TO = 1 as const;

  /**
   * Recipient is a copy recipient
   */
  public static readonly MAPI_CC = 2 as const;

  /**
   * Recipient is blind copy recipient
   */
  public static readonly MAPI_BCC = 3 as const;

  /**
   * Recipient is a P1 resend recipient
   */
  public static readonly MAPI_P1 = 0x10000000 as const;

  /**
   * Recipient is already processed
   */
  public static readonly MAPI_SUBMITTED = 0x80000000 as const;
}
