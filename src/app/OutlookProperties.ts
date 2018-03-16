// See PSTMessage.class for details on these properties
export const enum OutlookProperties {
    PSETID_Common = 1,
    PSETID_Address = 2,
    PSETID_Appointment = 4,
    PSETID_Meeting = 5,
    PSETID_Log = 6,

    PR_RTF_COMPRESSED = 0x1009,
    PR_NON_RECEIPT_NOTIFICATION_REQUESTED = 0x0C06,
    PR_ORIGINATOR_NON_DELIVERY_REPORT_REQUESTED = 0x0c08,
    PR_RECIPIENT_TYPE = 0x0c15,
    PR_MESSAGE_CODEPAGE = 0x3FFD,
    PR_INTERNET_CPID = 0x3FDE,
    PR_RTF_SYNC_BODY_CRC = 0x1006,
    PR_RTF_SYNC_BODY_COUNT = 0x1007,
    PR_RTF_SYNC_BODY_TAG = 0x1008,
    PR_RTF_SYNC_PREFIX_COUNT = 0x1010,
    PR_RTF_SYNC_TRAILING_COUNT = 0x1011,
    PR_BODY = 0x1000,
    PR_BODY_HTML = 0x1013,
    PR_IMPORTANCE = 0x0017,
    PR_MESSAGE_CLASS = 0x001A,
    PR_SUBJECT = 0x0037,
    PR_CLIENT_SUBMIT_TIME = 0x0039,
    PR_RECEIVED_BY_NAME = 0x0040,
    PR_SENT_REPRESENTING_NAME = 0x0042,
    PR_SENT_REPRESENTING_ADDRTYPE = 0x0064,
    PR_SENT_REPRESENTING_EMAIL_ADDRESS = 0x0065,
    PR_CONVERSATION_TOPIC = 0x0070,
    PR_RECEIVED_BY_ADDRTYPE = 0x0075,
    PR_RECEIVED_BY_EMAIL_ADDRESS = 0x0076,
    PR_TRANSPORT_MESSAGE_HEADERS = 0x007D,
    PR_MESSAGE_FLAGS = 0x0E07,
    PR_ORIGINATOR_DELIVERY_REPORT_REQUESTED = 0x0023,
    PR_PRIORITY = 0x0026,
    PR_READ_RECEIPT_REQUESTED = 0x0029,
    PR_RECIPIENT_REASSIGNMENT_PROHIBITED = 0x002B,
    PR_SENSITIVITY = 0x0036,
    PR_ORIGINAL_SENSITIVITY = 0x002E,
    PR_SENT_REPRESENTING_SEARCH_KEY = 0x003B,
    PR_RCVD_REPRESENTING_NAME = 0x0044,
    PR_ORIGINAL_SUBJECT = 0x0049,
    PR_REPLY_RECIPIENT_NAMES = 0x0050,
    PR_MESSAGE_TO_ME = 0x0057,
    PR_MESSAGE_CC_ME = 0x0058,
    PR_MESSAGE_RECIP_ME = 0x0059,
    PR_RESPONSE_REQUESTED = 0x0063,
    PR_ORIGINAL_DISPLAY_BCC = 0x0072,
    PR_ORIGINAL_DISPLAY_CC = 0x0073,
    PR_ORIGINAL_DISPLAY_TO = 0x0074,
    PR_RCVD_REPRESENTING_ADDRTYPE = 0x0077,
    PR_RCVD_REPRESENTING_EMAIL_ADDRESS = 0x0078,
    PR_REPLY_REQUESTED = 0x0C17,
    PR_SENDER_ENTRYID = 0x0C19,
    PR_SENDER_NAME = 0x0C1A,
    PR_SENDER_ADDRTYPE = 0x0C1E,
    PR_SENDER_EMAIL_ADDRESS = 0x0C1F,
    PR_MESSAGE_SIZE = 0x0E08,
    PR_INTERNET_ARTICLE_NUMBER = 0x0E23,
    PR_PRIMARY_SEND_ACCOUNT = 0x0E28,
    PR_NEXT_SEND_ACCT = 0x0E29,
    PR_OBJECT_TYPE = 0x0FFE,
    PR_DELETE_AFTER_SUBMIT = 0x0E01,
    PR_RESPONSIBILITY = 0x0E0F,
    PR_RTF_IN_SYNC = 0x0E1F,
    PR_DISPLAY_BCC = 0x0E02,
    PR_DISPLAY_CC = 0x0E03,
    PR_DISPLAY_TO = 0x0E04,
    PR_MESSAGE_DELIVERY_TIME = 0x0E06,
    PR_INTERNET_MESSAGE_ID = 0x1035,
    PR_IN_REPLY_TO_ID = 0x1042,
    PR_INTERNET_RETURN_PATH = 0x1046,
    PR_ICON_INDEX = 0x1080,
    PR_LAST_VERB_EXECUTED = 0x1081,
    PR_LAST_VERB_EXECUTION_TIME = 0x1082,
    PR_URL_COMP_NAME = 0x10F3,
    PR_ATTR_HIDDEN = 0x10f4,
    PR_EMAIL_ADDRESS = 0x3003,
    PR_ADDRTYPE = 0x3002,
    PR_COMMENT = 0x3004,
    PR_CREATION_TIME = 0x3007,
    PR_LAST_MODIFICATION_TIME = 0x3008,
    PR_ATTACH_DATA_BIN = 0x3701,
    PR_ATTACH_SIZE = 0x0E20,
    PR_ATTACH_FILENAME = 0x3704,
    PR_ATTACH_NUM = 0x0e21,
    PR_ATTACH_METHOD = 0x3705,
    PR_ATTACH_LONG_FILENAME = 0x3707,
    PR_ATTACH_PATHNAME = 0x3708,
    PR_RENDERING_POSITION = 0x370B,
    PR_ATTACH_LONG_PATHNAME = 0x370D,
    PR_ATTACH_MIME_TAG = 0x370E,
    PR_ATTACH_MIME_SEQUENCE = 0x3710,
    PR_ATTACH_CONTENT_ID = 0x3712,
    PR_ATTACH_FLAGS = 0x3714,
    PR_ACCOUNT = 0x3A00,
    PR_CALLBACK_TELEPHONE_NUMBER = 0x3A02,
    PR_GENERATION = 0x3A05,
    PR_GIVEN_NAME = 0x3A06,
    PR_GOVERNMENT_ID_NUMBER = 0x3A07,
    PR_BUSINESS_TELEPHONE_NUMBER = 0x3A08,
    PR_HOME_TELEPHONE_NUMBER = 0x3a09,
    PR_INITIALS = 0x3A0A,
    PR_KEYWORD = 0x3A0B,
    PR_LANGUAGE = 0x3A0C,
    PR_LOCATION = 0x3A0D,
    PR_MHS_COMMON_NAME = 0x3A0F,
    PR_ORGANIZATIONAL_ID_NUMBER = 0x3A10,
    PR_SURNAME = 0x3A11,
    PR_ORIGINAL_DISPLAY_NAME = 0x3A13,
    PR_POSTAL_ADDRESS = 0x3A15,
    PT_UNICODE = 0x3A16,
    PR_TITLE = 0x3A17,
    PR_DEPARTMENT_NAME = 0x3A18,
    PR_OFFICE_LOCATION = 0x3A19,
    PR_PRIMARY_TELEPHONE_NUMBER = 0x3A1A,
    PR_BUSINESS2_TELEPHONE_NUMBER = 0x3A1B,
    PR_MOBILE_TELEPHONE_NUMBER = 0x3A1C,
    PR_RADIO_TELEPHONE_NUMBER = 0x3A1D,
    PR_CAR_TELEPHONE_NUMBER = 0x3A1E,
    PR_OTHER_TELEPHONE_NUMBER = 0x3A1F,
    PR_TRANSMITABLE_DISPLAY_NAME = 0x3A20,
    PR_PAGER_TELEPHONE_NUMBER = 0x3A21,
    PR_PRIMARY_FAX_NUMBER = 0x3A23,
    PR_BUSINESS_FAX_NUMBER = 0x3A24,
    PR_HOME_FAX_NUMBER = 0x3A25,
    PR_COUNTRY = 0x3A26,
    PR_LOCALITY = 0x3A27,
    PR_STATE_OR_PROVINCE = 0x3A28,
    PR_STREET_ADDRESS = 0x3A29,
    PR_POSTAL_CODE = 0x3A2A,
    PR_POST_OFFICE_BOX = 0x3A2B,
    PR_TELEX_NUMBER = 0x3A2C,
    PR_ISDN_NUMBER = 0x3A2D,
    PR_ASSISTANT_TELEPHONE_NUMBER = 0x3A2E,
    PR_HOME2_TELEPHONE_NUMBER = 0x3A2F,
    PR_ASSISTANT = 0x3A30,
    PR_HOBBIES = 0x3A43,
    PR_MIDDLE_NAME = 0x3A44,
    PR_DISPLAY_NAME_PREFIX = 0x3A45,
    PR_PROFESSION = 0x3A46,
    PR_REFERRED_BY_NAME = 0x3A47,
    PR_SPOUSE_NAME = 0x3A48,
    PR_COMPUTER_NETWORK_NAME = 0x3A49,
    PR_CUSTOMER_ID = 0x3A4A,
    PR_TTYTDD_PHONE_NUMBER = 0x3A4B,
    PR_FTP_SITE = 0x3A4C,
    PR_MANAGER_NAME = 0x3A4E,
    PR_NICKNAME = 0x3A4F,
    PR_PERSONAL_HOME_PAGE = 0x3A50,
    PR_BUSINESS_HOME_PAGE = 0x3A51,
    PR_COMPANY_MAIN_PHONE_NUMBER = 0x3A57,
    PR_CHILDRENS_NAMES = 0x3A58,
    PR_HOME_ADDRESS_CITY = 0x3A59,
    PR_HOME_ADDRESS_COUNTRY = 0x3A5A,
    PR_HOME_ADDRESS_POSTAL_CODE = 0x3A5B,
    PR_HOME_ADDRESS_STATE_OR_PROVINCE = 0x3A5C,
    PR_HOME_ADDRESS_STREET = 0x3A5D,
    PR_HOME_ADDRESS_POST_OFFICE_BOX = 0x3A5E,
    PR_OTHER_ADDRESS_CITY = 0x3A5F,
    PR_OTHER_ADDRESS_COUNTRY = 0x3a60,
    PR_OTHER_ADDRESS_POSTAL_CODE = 0x3A61,
    PR_OTHER_ADDRESS_STATE_OR_PROVINCE = 0x3A62,
    PR_OTHER_ADDRESS_STREET = 0x3A63,
    PR_OTHER_ADDRESS_POST_OFFICE_BOX = 0x3A64,
    PR_FOLDER_TYPE = 0x3601,
    PR_CONTENT_COUNT = 0x3602,
    PR_CONTENT_UNREAD = 0x3603,
    PR_SUBFOLDERS = 0x360A,
    PR_CONTAINER_CLASS = 0x3613,
    PR_CONTAINER_FLAGS = 0x3600,
    PR_DISPLAY_NAME = 0x3001,
    PR_RECIPIENT_FLAGS = 0x5FFD,
    PR_SMTP_ADDRESS = 0x39FE,

    PidTagRecipientOrder = 0x5FDF,
    PidTagConversationId = 0x3013,
    PidTagConversationIndexTracking = 0x3016,
    PidLidLogType = 0x00008700,
    PidLidTaskStartDate = 0x00008104,
    PidLidTaskDueDate = 0x00008105,
    PidLidReminderSet = 0x00008503,
    PidLidReminderDelta = 0x00008501,
    PidLidLogStart = 0x00008706,
    PidLidLogDuration = 0x00008707,
    PidLidLogEnd = 0x00008708,
    PidLidLogFlags = 0x0000870C,
    PidLidLogDocumentPrinted = 0x0000870E,
    PidLidLogDocumentSaved = 0x0000870F,
    PidLidLogDocumentRouted = 0x00008710,
    PidLidLogDocumentPosted = 0x00008711,
    PidLidLogTypeDesc = 0x00008712,
    PidLidSendMeetingAsIcal = 0x00008200,
    PidLidBusyStatus = 0x00008205,
    PidLidLocation = 0x00008208,
    PidLidAppointmentStartWhole = 0x0000820D,
    PidLidAppointmentEndWhole = 0x0000820E,
    PidLidAppointmentDuration = 0x00008213,
    PidLidAppointmentColor = 0x00008214,
    PidLidAppointmentSubType = 0x00008215,
    PidLidAppointmentStateFlags = 0x00008217,
    PidLidResponseStatus = 0x00008218,
    PidLidRecurring = 0x00008223,
    PidLidExceptionReplaceTime = 0x00008228,
    PidLidRecurrenceType = 0x00008231,
    PidLidRecurrencePattern = 0x00008232,
    PidLidAppointmentRecur = 0x00008216,
    PidLidTimeZoneStruct = 0x00008233,
    PidLidAllAttendeesString = 0x00008238,
    PidLidToAttendeesString = 0x0000823B,
    PidLidCcAttendeesString = 0x0000823C,
    PidLidAppointmentSequence = 0x00008201,
    PidLidConferencingCheck = 0x00008240,
    PidLidConferencingType = 0x00008241,
    PidLidDirectory = 0x00008242,
    PidLidOrganizerAlias = 0x00008243,
    PidLidNetShowUrl = 0x00008248,
    PidLidCollaborateDoc = 0x00008247,
    PidLidAttendeeCriticalChange = 0x00000001,
    PidLidAppointmentCounterProposal = 0x00008257,
    PidLidIsSilent = 0x00000004,
    PidLidRequiredAttendees = 0x00000006,
    PidTagMessageLocaleId = 0x3FF1,
    PidLidFileUnder = 0x00008005,
    PidLidHomeAddress = 0x0000801A,
    PidLidWorkAddress = 0x0000801B,
    PidLidOtherAddress = 0x0000801C,
    PidLidPostalAddressId = 0x00008022,
    PidLidHtml = 0x0000802B,
    PidLidWorkAddressStreet = 0x00008045,
    PidLidWorkAddressCity = 0x00008046,
    PidLidWorkAddressState = 0x00008047,
    PidLidWorkAddressPostalCode = 0x00008048,
    PidLidWorkAddressCountry = 0x00008049,
    PidLidWorkAddressPostOfficeBox = 0x0000804A,
    PidLidInstantMessagingAddress = 0x00008062,
    PidLidEmail1DisplayName = 0x00008080,
    PidLidEmail1AddressType = 0x00008082,
    PidLidEmail1EmailAddress = 0x00008083,
    PidLidEmail1OriginalDisplayName = 0x00008084,
    PidLidEmail2DisplayName = 0x00008090,
    PidLidEmail2AddressType = 0x00008092,
    PidLidEmail2EmailAddress = 0x00008093,
    PidLidEmail2OriginalDisplayName = 0x00008094,
    PidLidEmail3DisplayName = 0x000080A0,
    PidLidEmail3AddressType = 0x000080A2,
    PidLidEmail3EmailAddress = 0x000080A3,
    PidLidEmail3OriginalDisplayName = 0x000080A4,
    PidLidFax1AddressType = 0x000080B2,
    PidLidFax1EmailAddress = 0x000080B3,
    PidLidFax1OriginalDisplayName = 0x000080B4,
    PidLidFax2AddressType = 0x000080C2,
    PidLidFax2EmailAddress = 0x000080C3,
    PidLidFax2OriginalDisplayName = 0x000080C4,
    PidLidFax3AddressType = 0x000080D2,
    PidLidFax3EmailAddress = 0x000080D3,
    PidLidFax3OriginalDisplayName = 0x000080D4,
    PidLidFreeBusyLocation = 0x000080D8,
    PidTagBirthday = 0x3A42,
    PidTagWeddingAnniversary = 0x3A41,
    PidLidPercentComplete = 0x00008102,
    PidLidTaskStatus = 0x00008101,  
    PidLidTaskDateCompleted = 0x0000810F,
    PidLidTaskActualEffort = 0x00008110,
    PidLidTaskEstimatedEffort = 0x00008111,
    PidLidTaskVersion = 0x00008112,
    PidLidTaskComplete = 0x0000811C,
    PidLidTaskOwner = 0x0000811F,
    PidLidTaskAssigner = 0x00008121,
    PidLidTaskLastUser = 0x00008122,
    PidLidTaskOrdinal = 0x00008123,
    PidLidTaskFRecurring = 0x00008126,
    PidLidTaskOwnership = 0x00008129,
    PidLidTaskAcceptanceState = 0x0000812A
}
