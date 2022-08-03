export declare enum OutlookProperties {
    PSETID_Common = 1,
    PSETID_Address = 2,
    PSETID_Appointment = 4,
    PSETID_Meeting = 5,
    PSETID_Log = 6,
    PR_RTF_COMPRESSED = 4105,
    PR_NON_RECEIPT_NOTIFICATION_REQUESTED = 3078,
    PR_ORIGINATOR_NON_DELIVERY_REPORT_REQUESTED = 3080,
    PR_RECIPIENT_TYPE = 3093,
    PR_MESSAGE_CODEPAGE = 16381,
    PR_INTERNET_CPID = 16350,
    PR_RTF_SYNC_BODY_CRC = 4102,
    PR_RTF_SYNC_BODY_COUNT = 4103,
    PR_RTF_SYNC_BODY_TAG = 4104,
    PR_RTF_SYNC_PREFIX_COUNT = 4112,
    PR_RTF_SYNC_TRAILING_COUNT = 4113,
    PR_BODY = 4096,
    PR_BODY_HTML = 4115,
    PR_IMPORTANCE = 23,
    PR_MESSAGE_CLASS = 26,
    PR_SUBJECT = 55,
    PR_CLIENT_SUBMIT_TIME = 57,
    PR_RECEIVED_BY_NAME = 64,
    PR_SENT_REPRESENTING_NAME = 66,
    PR_SENT_REPRESENTING_ADDRTYPE = 100,
    PR_SENT_REPRESENTING_EMAIL_ADDRESS = 101,
    PR_CONVERSATION_TOPIC = 112,
    PR_RECEIVED_BY_ADDRTYPE = 117,
    PR_RECEIVED_BY_EMAIL_ADDRESS = 118,
    PR_TRANSPORT_MESSAGE_HEADERS = 125,
    PR_MESSAGE_FLAGS = 3591,
    PR_ORIGINATOR_DELIVERY_REPORT_REQUESTED = 35,
    PR_PRIORITY = 38,
    PR_READ_RECEIPT_REQUESTED = 41,
    PR_RECIPIENT_REASSIGNMENT_PROHIBITED = 43,
    PR_SENSITIVITY = 54,
    PR_ORIGINAL_SENSITIVITY = 46,
    PR_SENT_REPRESENTING_SEARCH_KEY = 59,
    PR_RCVD_REPRESENTING_NAME = 68,
    PR_ORIGINAL_SUBJECT = 73,
    PR_REPLY_RECIPIENT_NAMES = 80,
    PR_MESSAGE_TO_ME = 87,
    PR_MESSAGE_CC_ME = 88,
    PR_MESSAGE_RECIP_ME = 89,
    PR_RESPONSE_REQUESTED = 99,
    PR_ORIGINAL_DISPLAY_BCC = 114,
    PR_ORIGINAL_DISPLAY_CC = 115,
    PR_ORIGINAL_DISPLAY_TO = 116,
    PR_RCVD_REPRESENTING_ADDRTYPE = 119,
    PR_RCVD_REPRESENTING_EMAIL_ADDRESS = 120,
    PR_REPLY_REQUESTED = 3095,
    PR_SENDER_ENTRYID = 3097,
    PR_SENDER_NAME = 3098,
    PR_SENDER_ADDRTYPE = 3102,
    PR_SENDER_EMAIL_ADDRESS = 3103,
    PR_MESSAGE_SIZE = 3592,
    PR_INTERNET_ARTICLE_NUMBER = 3619,
    PR_PRIMARY_SEND_ACCOUNT = 3624,
    PR_NEXT_SEND_ACCT = 3625,
    PR_OBJECT_TYPE = 4094,
    PR_DELETE_AFTER_SUBMIT = 3585,
    PR_RESPONSIBILITY = 3599,
    PR_RTF_IN_SYNC = 3615,
    PR_DISPLAY_BCC = 3586,
    PR_DISPLAY_CC = 3587,
    PR_DISPLAY_TO = 3588,
    PR_MESSAGE_DELIVERY_TIME = 3590,
    PR_INTERNET_MESSAGE_ID = 4149,
    PR_IN_REPLY_TO_ID = 4162,
    PR_INTERNET_RETURN_PATH = 4166,
    PR_ICON_INDEX = 4224,
    PR_LAST_VERB_EXECUTED = 4225,
    PR_LAST_VERB_EXECUTION_TIME = 4226,
    PR_URL_COMP_NAME = 4339,
    PR_ATTR_HIDDEN = 4340,
    PR_EMAIL_ADDRESS = 12291,
    PR_ADDRTYPE = 12290,
    PR_COMMENT = 12292,
    PR_CREATION_TIME = 12295,
    PR_LAST_MODIFICATION_TIME = 12296,
    PR_ATTACH_DATA_BIN = 14081,
    PR_ATTACH_SIZE = 3616,
    PR_ATTACH_FILENAME = 14084,
    PR_ATTACH_NUM = 3617,
    PR_ATTACH_METHOD = 14085,
    PR_ATTACH_LONG_FILENAME = 14087,
    PR_ATTACH_PATHNAME = 14088,
    PR_RENDERING_POSITION = 14091,
    PR_ATTACH_LONG_PATHNAME = 14093,
    PR_ATTACH_MIME_TAG = 14094,
    PR_ATTACH_MIME_SEQUENCE = 14096,
    PR_ATTACH_CONTENT_ID = 14098,
    PR_ATTACH_FLAGS = 14100,
    PR_ACCOUNT = 14848,
    PR_CALLBACK_TELEPHONE_NUMBER = 14850,
    PR_GENERATION = 14853,
    PR_GIVEN_NAME = 14854,
    PR_GOVERNMENT_ID_NUMBER = 14855,
    PR_BUSINESS_TELEPHONE_NUMBER = 14856,
    PR_HOME_TELEPHONE_NUMBER = 14857,
    PR_INITIALS = 14858,
    PR_KEYWORD = 14859,
    PR_LANGUAGE = 14860,
    PR_LOCATION = 14861,
    PR_MHS_COMMON_NAME = 14863,
    PR_ORGANIZATIONAL_ID_NUMBER = 14864,
    PR_SURNAME = 14865,
    PR_ORIGINAL_DISPLAY_NAME = 14867,
    PR_POSTAL_ADDRESS = 14869,
    PT_UNICODE = 14870,
    PR_TITLE = 14871,
    PR_DEPARTMENT_NAME = 14872,
    PR_OFFICE_LOCATION = 14873,
    PR_PRIMARY_TELEPHONE_NUMBER = 14874,
    PR_BUSINESS2_TELEPHONE_NUMBER = 14875,
    PR_MOBILE_TELEPHONE_NUMBER = 14876,
    PR_RADIO_TELEPHONE_NUMBER = 14877,
    PR_CAR_TELEPHONE_NUMBER = 14878,
    PR_OTHER_TELEPHONE_NUMBER = 14879,
    PR_TRANSMITABLE_DISPLAY_NAME = 14880,
    PR_PAGER_TELEPHONE_NUMBER = 14881,
    PR_PRIMARY_FAX_NUMBER = 14883,
    PR_BUSINESS_FAX_NUMBER = 14884,
    PR_HOME_FAX_NUMBER = 14885,
    PR_COUNTRY = 14886,
    PR_LOCALITY = 14887,
    PR_STATE_OR_PROVINCE = 14888,
    PR_STREET_ADDRESS = 14889,
    PR_POSTAL_CODE = 14890,
    PR_POST_OFFICE_BOX = 14891,
    PR_TELEX_NUMBER = 14892,
    PR_ISDN_NUMBER = 14893,
    PR_ASSISTANT_TELEPHONE_NUMBER = 14894,
    PR_HOME2_TELEPHONE_NUMBER = 14895,
    PR_ASSISTANT = 14896,
    PR_HOBBIES = 14915,
    PR_MIDDLE_NAME = 14916,
    PR_DISPLAY_NAME_PREFIX = 14917,
    PR_PROFESSION = 14918,
    PR_REFERRED_BY_NAME = 14919,
    PR_SPOUSE_NAME = 14920,
    PR_COMPUTER_NETWORK_NAME = 14921,
    PR_CUSTOMER_ID = 14922,
    PR_TTYTDD_PHONE_NUMBER = 14923,
    PR_FTP_SITE = 14924,
    PR_MANAGER_NAME = 14926,
    PR_NICKNAME = 14927,
    PR_PERSONAL_HOME_PAGE = 14928,
    PR_BUSINESS_HOME_PAGE = 14929,
    PR_COMPANY_MAIN_PHONE_NUMBER = 14935,
    PR_CHILDRENS_NAMES = 14936,
    PR_HOME_ADDRESS_CITY = 14937,
    PR_HOME_ADDRESS_COUNTRY = 14938,
    PR_HOME_ADDRESS_POSTAL_CODE = 14939,
    PR_HOME_ADDRESS_STATE_OR_PROVINCE = 14940,
    PR_HOME_ADDRESS_STREET = 14941,
    PR_HOME_ADDRESS_POST_OFFICE_BOX = 14942,
    PR_OTHER_ADDRESS_CITY = 14943,
    PR_OTHER_ADDRESS_COUNTRY = 14944,
    PR_OTHER_ADDRESS_POSTAL_CODE = 14945,
    PR_OTHER_ADDRESS_STATE_OR_PROVINCE = 14946,
    PR_OTHER_ADDRESS_STREET = 14947,
    PR_OTHER_ADDRESS_POST_OFFICE_BOX = 14948,
    PR_FOLDER_TYPE = 13825,
    PR_CONTENT_COUNT = 13826,
    PR_CONTENT_UNREAD = 13827,
    PR_SUBFOLDERS = 13834,
    PR_CONTAINER_CLASS = 13843,
    PR_CONTAINER_FLAGS = 13824,
    PR_DISPLAY_NAME = 12289,
    PR_RECIPIENT_FLAGS = 24573,
    PR_SMTP_ADDRESS = 14846,
    PidTagRecipientOrder = 24543,
    PidTagConversationId = 12307,
    PidTagConversationIndexTracking = 12310,
    PidLidLogType = 34560,
    PidLidTaskStartDate = 33028,
    PidLidTaskDueDate = 33029,
    PidLidReminderSet = 34051,
    PidLidReminderDelta = 34049,
    PidLidLogStart = 34566,
    PidLidLogDuration = 34567,
    PidLidLogEnd = 34568,
    PidLidLogFlags = 34572,
    PidLidLogDocumentPrinted = 34574,
    PidLidLogDocumentSaved = 34575,
    PidLidLogDocumentRouted = 34576,
    PidLidLogDocumentPosted = 34577,
    PidLidLogTypeDesc = 34578,
    PidLidSendMeetingAsIcal = 33280,
    PidLidBusyStatus = 33285,
    PidLidLocation = 33288,
    PidLidAppointmentStartWhole = 33293,
    PidLidAppointmentEndWhole = 33294,
    PidLidAppointmentDuration = 33299,
    PidLidAppointmentColor = 33300,
    PidLidAppointmentSubType = 33301,
    PidLidAppointmentStateFlags = 33303,
    PidLidResponseStatus = 33304,
    PidLidRecurring = 33315,
    PidLidExceptionReplaceTime = 33320,
    PidLidRecurrenceType = 33329,
    PidLidRecurrencePattern = 33330,
    PidLidAppointmentRecur = 33302,
    PidLidTimeZoneStruct = 33331,
    PidLidAllAttendeesString = 33336,
    PidLidToAttendeesString = 33339,
    PidLidCcAttendeesString = 33340,
    PidLidAppointmentSequence = 33281,
    PidLidConferencingCheck = 33344,
    PidLidConferencingType = 33345,
    PidLidDirectory = 33346,
    PidLidOrganizerAlias = 33347,
    PidLidNetShowUrl = 33352,
    PidLidCollaborateDoc = 33351,
    PidLidAttendeeCriticalChange = 1,
    PidLidAppointmentCounterProposal = 33367,
    PidLidIsSilent = 4,
    PidLidRequiredAttendees = 6,
    PidTagMessageLocaleId = 16369,
    PidLidFileUnder = 32773,
    PidLidHomeAddress = 32794,
    PidLidWorkAddress = 32795,
    PidLidOtherAddress = 32796,
    PidLidPostalAddressId = 32802,
    PidLidHtml = 32811,
    PidLidWorkAddressStreet = 32837,
    PidLidWorkAddressCity = 32838,
    PidLidWorkAddressState = 32839,
    PidLidWorkAddressPostalCode = 32840,
    PidLidWorkAddressCountry = 32841,
    PidLidWorkAddressPostOfficeBox = 32842,
    PidLidInstantMessagingAddress = 32866,
    PidLidEmail1DisplayName = 32896,
    PidLidEmail1AddressType = 32898,
    PidLidEmail1EmailAddress = 32899,
    PidLidEmail1OriginalDisplayName = 32900,
    PidLidEmail2DisplayName = 32912,
    PidLidEmail2AddressType = 32914,
    PidLidEmail2EmailAddress = 32915,
    PidLidEmail2OriginalDisplayName = 32916,
    PidLidEmail3DisplayName = 32928,
    PidLidEmail3AddressType = 32930,
    PidLidEmail3EmailAddress = 32931,
    PidLidEmail3OriginalDisplayName = 32932,
    PidLidFax1AddressType = 32946,
    PidLidFax1EmailAddress = 32947,
    PidLidFax1OriginalDisplayName = 32948,
    PidLidFax2AddressType = 32962,
    PidLidFax2EmailAddress = 32963,
    PidLidFax2OriginalDisplayName = 32964,
    PidLidFax3AddressType = 32978,
    PidLidFax3EmailAddress = 32979,
    PidLidFax3OriginalDisplayName = 32980,
    PidLidFreeBusyLocation = 32984,
    PidTagBirthday = 14914,
    PidTagWeddingAnniversary = 14913,
    PidLidPercentComplete = 33026,
    PidLidTaskStatus = 33025,
    PidLidTaskDeadOccurrence = 33033,
    PidLidTaskDateCompleted = 33039,
    PidLidTaskActualEffort = 33040,
    PidLidTaskEstimatedEffort = 33041,
    PidLidTaskVersion = 33042,
    PidLidTaskRecurrence = 33046,
    PidLidTaskComplete = 33052,
    PidLidTaskOwner = 33055,
    PidLidTaskAssigner = 33057,
    PidLidTaskLastUser = 33058,
    PidLidTaskOrdinal = 33059,
    PidLidTaskFRecurring = 33062,
    PidLidTaskOwnership = 33065,
    PidLidTaskAcceptanceState = 33066,
    PidLidYomiLastName = 32813,
    PidLidYomiFirstName = 32812,
    PidLidYomiCompanyName = 32814
}
