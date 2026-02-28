export interface IPSTContact {
  /**
   * Contains the recipient's account name.
   * https://msdn.microsoft.com/en-us/library/office/cc842401.aspx
   * @readonly
   */
  get account(): string;

  /**
   * Contains a telephone number that the message recipient can use to reach the sender.
   * https://msdn.microsoft.com/en-us/library/office/cc839943.aspx
   * @readonly
   */
  get callbackTelephoneNumber(): string;

  /**
   * Contains a generational abbreviation that follows the full name of the recipient.
   * https://msdn.microsoft.com/en-us/library/office/cc842136.aspx
   * @readonly
   */
  get generation(): string;

  /**
   * Contains the first or given name of the recipient.
   * https://msdn.microsoft.com/en-us/library/office/cc815351.aspx
   * @readonly
   */
  get givenName(): string;

  /**
   * Contains a government identifier for the recipient.
   * https://msdn.microsoft.com/en-us/library/office/cc815890.aspx
   * @readonly
   */
  get governmentIdNumber(): string;

  /**
   * Contains the primary telephone number of the recipient's place of business.
   * https://msdn.microsoft.com/en-us/library/office/cc839937.aspx
   * @readonly
   */
  get businessTelephoneNumber(): string;

  /**
   * Contains the primary telephone number of the recipient's home.
   * https://msdn.microsoft.com/en-us/library/office/cc815389.aspx
   * @readonly
   */
  get homeTelephoneNumber(): string;

  /**
   * Contains the initials for parts of the full name of the recipient.
   * https://msdn.microsoft.com/en-us/library/office/cc839843.aspx
   * @readonly
   */
  get initials(): string;

  /**
   * Contains a keyword that identifies the recipient to the recipient's system administrator.
   * https://msdn.microsoft.com/en-us/library/office/cc842250.aspx
   * @readonly
   */
  get keyword(): string;

  /**
   * Contains a value that indicates the language in which the messaging user is writing messages.
   * https://msdn.microsoft.com/en-us/library/office/cc839724.aspx
   * @readonly
   */
  get language(): string;

  /**
   * Contains the location of the recipient in a format that is useful to the recipient's organization.
   * https://msdn.microsoft.com/en-us/library/office/cc815567.aspx
   * @readonly
   */
  get location(): string;

  /**
   * Contains the common name of the message handling system.
   * https://msdn.microsoft.com/en-us/library/office/cc842474.aspx
   * @readonly
   */
  get mhsCommonName(): string;

  /**
   * Contains an organizational ID number for the contact, such as an employee ID number.
   * https://msdn.microsoft.com/en-us/library/office/cc765672.aspx
   * @readonly
   */
  get organizationalIdNumber(): string;

  /**
   * Contains the last or surname of the recipient.
   * https://msdn.microsoft.com/en-us/library/office/cc765704.aspx
   * @readonly
   */
  get surname(): string;

  /**
   * Contains the original display name for an entry copied from an address book to a personal address book or other writable address book.
   * https://msdn.microsoft.com/en-us/library/office/cc765709.aspx
   * @readonly
   */
  get originalDisplayName(): string;

  /**
   * Contains the recipient's postal address.
   * https://msdn.microsoft.com/en-us/library/office/cc842549.aspx
   * @readonly
   */
  get postalAddress(): string;

  /**
   * Contains the recipient's company name.
   * https://msdn.microsoft.com/en-us/library/office/cc842192.aspx
   * @readonly
   */
  get companyName(): string;

  /**
   * Contains the recipient's job title.
   * https://msdn.microsoft.com/en-us/library/office/cc815831.aspx
   * @readonly
   */
  get title(): string;

  /**
   * Contains a name for the department in which the recipient works.
   * https://msdn.microsoft.com/en-us/library/office/cc839825.aspx
   * @readonly
   */
  get departmentName(): string;

  /**
   * Contains the recipient's office location.
   * https://msdn.microsoft.com/en-us/library/office/cc842269.aspx
   * @readonly
   */
  get officeLocation(): string;

  /**
   * Contains the recipient's primary telephone number.
   * https://msdn.microsoft.com/en-us/library/office/cc839969.aspx
   * @readonly
   */
  get primaryTelephoneNumber(): string;

  /**
   * Contains a secondary telephone number at the recipient's place of business.
   * https://msdn.microsoft.com/en-us/library/office/cc841990.aspx
   * @readonly
   */
  get business2TelephoneNumber(): string;

  /**
   * Contains the recipient's cellular telephone number.
   * https://msdn.microsoft.com/en-us/library/office/cc839798.aspx
   * @readonly
   */
  get mobileTelephoneNumber(): string;

  /**
   * Contains the recipient's radio telephone number.
   * https://msdn.microsoft.com/en-us/library/office/cc839806.aspx
   * @readonly
   */
  get radioTelephoneNumber(): string;

  /**
   * Contains the recipient's car telephone number.
   * https://msdn.microsoft.com/en-us/library/office/cc815394.aspx
   * @readonly
   */
  get carTelephoneNumber(): string;

  /**
   * Contains an alternate telephone number for the recipient.
   * https://msdn.microsoft.com/en-us/library/office/cc839561.aspx
   * @readonly
   */
  get otherTelephoneNumber(): string;

  /**
   * Contains a recipient's display name in a secure form that cannot be changed.
   * https://msdn.microsoft.com/en-us/library/office/cc815723.aspx
   * @readonly
   */
  get transmittableDisplayName(): string;

  /**
   * Contains the recipient's pager telephone number.
   * https://msdn.microsoft.com/en-us/library/office/cc765824.aspx
   * @readonly
   */
  get pagerTelephoneNumber(): string;

  /**
   * Contains the telephone number of the recipient's primary fax machine.
   * https://msdn.microsoft.com/en-us/library/office/cc815713.aspx
   * @readonly
   */
  get primaryFaxNumber(): string;

  /**
   * Contains the telephone number of the recipient's business fax machine.
   * https://msdn.microsoft.com/en-us/library/office/cc765799.aspx
   * @readonly
   */
  get businessFaxNumber(): string;

  /**
   * Contains the telephone number of the recipient's home fax machine.
   * https://msdn.microsoft.com/en-us/library/office/cc842109.aspx
   * @readonly
   */
  get homeFaxNumber(): string;

  /**
   * Contains the name of the recipient's country/region.
   * https://msdn.microsoft.com/en-us/library/office/cc842494.aspx
   * @readonly
   */
  get businessAddressCountry(): string;

  /**
   * Contains the name of the recipient's locality, such as the town or city.
   * https://msdn.microsoft.com/en-us/library/office/cc815711.aspx
   * @readonly
   */
  get businessAddressCity(): string;

  /**
   * Contains the name of the recipient's state or province.
   * https://msdn.microsoft.com/en-us/library/office/cc839544.aspx
   * @readonly
   */
  get businessAddressStateOrProvince(): string;

  /**
   * Contains the recipient's street address.
   * https://msdn.microsoft.com/en-us/library/office/cc765810.aspx
   * @readonly
   */
  get businessAddressStreet(): string;

  /**
   * Contains the postal code for the recipient's postal address.
   * https://msdn.microsoft.com/en-us/library/office/cc839851.aspx
   * @readonly
   */
  get businessPostalCode(): string;

  /**
   * Contains the number or identifier of the recipient's post office box.
   * https://msdn.microsoft.com/en-us/library/office/cc815522.aspx
   * @readonly
   */
  get businessPoBox(): string;

  /**
   * Contains the recipient's telex number.
   * https://msdn.microsoft.com/en-us/library/office/cc765894.aspx
   * @readonly
   */
  get telexNumber(): string;

  /**
   * Contains the recipient's ISDN-capable telephone number.
   * https://msdn.microsoft.com/en-us/library/office/cc765863.aspx
   * @readonly
   */
  get isdnNumber(): string;

  /**
   * Contains the telephone number of the recipient's administrative assistant.
   * https://msdn.microsoft.com/en-us/library/office/cc840012.aspx
   * @readonly
   */
  get assistantTelephoneNumber(): string;

  /**
   * Contains a secondary telephone number at the recipient's home.
   * https://msdn.microsoft.com/en-us/library/office/cc815540.aspx
   * @readonly
   */
  get home2TelephoneNumber(): string;

  /**
   * Contains the name of the recipient's administrative assistant.
   * https://msdn.microsoft.com/en-us/library/office/cc815319.aspx
   * @readonly
   */
  get assistant(): string;

  /**
   * Contains the names of the hobbies of the messaging user.
   * https://msdn.microsoft.com/en-us/library/office/cc815391.aspx
   * @readonly
   */
  get hobbies(): string;

  /**
   * Contains the middle name of a contact.
   * https://msdn.microsoft.com/en-us/library/office/cc815329.aspx
   * @readonly
   */
  get middleName(): string;

  /**
   * Contains the display name prefix (such as Miss, Mr., Mrs.) for the messaging user.
   * https://msdn.microsoft.com/en-us/library/office/cc765538.aspx
   * @readonly
   */
  get displayNamePrefix(): string;

  /**
   * Contains the profession of the user.
   * https://msdn.microsoft.com/en-us/library/office/cc765792.aspx
   * @readonly
   */
  get profession(): string;

  /**
   * Contains the name of the mail user's referral.
   * https://msdn.microsoft.com/en-us/library/office/cc765803.aspx
   * @readonly
   */
  get preferredByName(): string;

  /**
   * Contains the user’s spouse name.
   * https://msdn.microsoft.com/en-us/library/office/cc765832.aspx
   * @readonly
   */
  get spouseName(): string;

  /**
   * Contains the name of the network used to transmit the message.
   * https://msdn.microsoft.com/en-us/library/office/cc839633.aspx
   * @readonly
   */
  get computerNetworkName(): string;

  /**
   * Contains the contact’s customer ID number.
   * https://msdn.microsoft.com/en-us/library/office/cc842178.aspx
   * @readonly
   */
  get customerId(): string;

  /**
   * Contains the telephone number for the contact’s text telephone (TTY) or telecommunication device for the deaf (TDD).
   * https://msdn.microsoft.com/en-us/library/office/cc765580.aspx
   * @readonly
   */
  get ttytddPhoneNumber(): string;

  /**
   * Contains the contact’s File Transfer Protocol (FTP) URL. FTP is a protocol that is used to transfer data, as specified in [RFC959].
   * https://msdn.microsoft.com/en-us/library/office/cc839830.aspx
   * @readonly
   */
  get ftpSite(): string;

  /**
   * Contains the name of the recipient's manager.
   * https://msdn.microsoft.com/en-us/library/office/cc842009.aspx
   * @readonly
   */
  get managerName(): string;

  /**
   * Contains the nickname of the contact.
   * https://msdn.microsoft.com/en-us/library/office/cc765603.aspx
   * @readonly
   */
  get nickname(): string;

  /**
   * Contains the URL of a user's personal home page.
   * https://msdn.microsoft.com/en-us/library/office/cc765751.aspx
   * @readonly
   */
  get personalHomePage(): string;

  /**
   * Contains the URL of the home page for the business.
   * https://msdn.microsoft.com/en-us/library/office/cc842385.aspx
   * @readonly
   */
  get businessHomePage(): string;

  /**
   * Get the note associated with the contact.
   * @readonly
   */
  get note(): string;

  /**
   * Get a named string item from the map
   */
  getNamedStringItem(key: number): string;

  /**
   * Contains the main telephone number for a company
   * https://msdn.microsoft.com/en-us/library/office/cc839651.aspx
   * @readonly
   */
  get companyMainPhoneNumber(): string;

  /**
   * Contains a list of names of children
   * https://msdn.microsoft.com/en-us/library/office/cc839533.aspx
   * @readonly
   */
  get childrensNames(): string;

  /**
   * Contains the city for the recipient's home address.
   * https://msdn.microsoft.com/en-us/library/office/cc815582.aspx
   * @readonly
   */
  get homeAddressCity(): string;

  /**
   * Contains the county in a contact's address.
   * https://msdn.microsoft.com/en-us/library/office/cc842548.aspx
   * @readonly
   */
  get homeAddressCountry(): string;

  /**
   * Contains the postal code for the user's home address.
   * https://msdn.microsoft.com/en-us/library/office/cc815880.aspx
   * @readonly
   */
  get homeAddressPostalCode(): string;

  /**
   * Contains the state or province portion of a user's address.
   * https://msdn.microsoft.com/en-us/library/office/cc839958.aspx
   * @readonly
   */
  get homeAddressStateOrProvince(): string;

  /**
   * Contains the street portion of a user's address.
   * https://msdn.microsoft.com/en-us/library/office/cc841997.aspx
   * @readonly
   */
  get homeAddressStreet(): string;

  /**
   * Contains the post office box information for a user's address.
   * https://msdn.microsoft.com/en-us/library/office/cc842440.aspx
   * @readonly
   */
  get homeAddressPostOfficeBox(): string;

  /**
   * Contains the name of the mail user's other locality, such as the town or city.
   * https://msdn.microsoft.com/en-us/library/office/cc765881.aspx
   * @readonly
   */
  get otherAddressCity(): string;

  /**
   * Contains the mail user's other country/region.
   * https://msdn.microsoft.com/en-us/library/office/cc765814.aspx
   * @readonly
   */
  get otherAddressCountry(): string;
  /**
   * Contains the postal code for the mail user's other postal address.
   * https://msdn.microsoft.com/en-us/library/office/cc842261.aspx
   * @readonly
   */
  get otherAddressPostalCode(): string;

  /**
   * Contains the name of state or province used in the other address.
   * https://msdn.microsoft.com/en-us/library/office/cc815782.aspx
   * @readonly
   */
  get otherAddressStateOrProvince(): string;

  /**
   * Contains the mail user's other street address.
   * https://msdn.microsoft.com/en-us/library/office/cc839546.aspx
   * @readonly
   */
  get otherAddressStreet(): string;

  /**
   * Contains the post office box for a contact's other address.
   * https://msdn.microsoft.com/en-us/library/office/cc842396.aspx
   * @readonly
   */
  get otherAddressPostOfficeBox(): string;

  /**
   * Specifies the name under which the contact is filed when displaying a list of contacts.
   * https://msdn.microsoft.com/en-us/library/office/cc842002.aspx
   * @readonly
   */
  get fileUnder(): string;

  /**
   * Specifies the complete address of the contact’s home address.
   * https://msdn.microsoft.com/en-us/library/office/cc839539.aspx
   * @readonly
   */
  get homeAddress(): string;

  /**
   * Specifies the contact's complete work address.
   * https://msdn.microsoft.com/en-us/library/office/cc815905.aspx
   * @readonly
   */
  get workAddress(): string;

  /**
   * Specifies the complete address of the contact’s other address.
   * https://msdn.microsoft.com/en-us/library/office/cc815383.aspx
   * @readonly
   */
  get otherAddress(): string;

  /**
   * Specifies which physical address is the contact’s mailing address.
   * https://msdn.microsoft.com/en-us/library/office/cc815430.aspx
   * @readonly
   */
  get postalAddressId(): number;

  /**
   * Specifies the contact’s business Web page URL.
   * https://msdn.microsoft.com/en-us/library/office/cc842001.aspx
   * @readonly
   */
  get html(): string;

  /**
   * Specifies the street portion of the contact's work mailing address.
   * https://msdn.microsoft.com/en-us/library/office/cc815537.aspx
   * @readonly
   */
  get workAddressStreet(): string;

  /**
   * Specifies the city or locality portion of the contact's work address.
   * https://msdn.microsoft.com/en-us/library/office/cc765923.aspx
   * @readonly
   */
  get workAddressCity(): string;

  /**
   * Specifies the state or province portion of the contact's work mailing address.
   * https://msdn.microsoft.com/en-us/library/office/cc842152.aspx
   * @readonly
   */
  get workAddressState(): string;

  /**
   * Specifies the postal code (ZIP code) portion of the contact's work address.
   * https://msdn.microsoft.com/en-us/library/office/cc842066.aspx
   * @readonly
   */
  get workAddressPostalCode(): string;

  /**
   * Specifies the country or region portion of the contact's work address.
   * https://msdn.microsoft.com/en-us/library/office/cc765698.aspx
   * @readonly
   */
  get workAddressCountry(): string;

  /**
   * Specifies the post office box portion of the contact's work.
   * https://msdn.microsoft.com/en-us/library/office/cc815563.aspx
   * @readonly
   */
  get workAddressPostOfficeBox(): string;

  /**
   * Specifies the contact’s instant messaging address.
   * https://msdn.microsoft.com/en-us/library/office/cc815607.aspx
   * @readonly
   */
  get instantMessagingAddress(): string;

  /**
   * Specifies the user-readable display name for the first e-mail address.
   * https://msdn.microsoft.com/en-us/library/office/cc815460.aspx
   * @readonly
   */
  get email1DisplayName(): string;

  /**
   * Specifies the address type of the first e-mail address.
   * https://msdn.microsoft.com/en-us/library/office/cc815570.aspx
   * @readonly
   */
  get email1AddressType(): string;

  /**
   * Specifies the first e-mail address of the contact.
   * https://msdn.microsoft.com/en-us/library/office/cc842050.aspx
   * @readonly
   */
  get email1EmailAddress(): string;

  /**
   * Specifies the first display name that corresponds to the e-mail address that is specified for the contact.
   * https://msdn.microsoft.com/en-us/library/office/cc815564.aspx
   * @readonly
   */
  get email1OriginalDisplayName(): string;

  /**
   * Specifies the user-readable display name for the second e-mail address.
   * https://msdn.microsoft.com/en-us/library/office/cc839675.aspx
   * @readonly
   */
  get email2DisplayName(): string;

  /**
   * Specifies the address type of the second e-mail address.
   * https://msdn.microsoft.com/en-us/library/office/cc815361.aspx
   * @readonly
   */
  get email2AddressType(): string;

  /**
   * Specifies the second e-mail address of the contact.
   * https://msdn.microsoft.com/en-us/library/office/cc842205.aspx
   * @readonly
   */
  get email2EmailAddress(): string;

  /**
   * Specifies the second display name that corresponds to the e-mail address specified for the contact.
   * https://msdn.microsoft.com/en-us/library/office/cc765618.aspx
   * @readonly
   */
  get email2OriginalDisplayName(): string;

  /**
   * Specifies the user-readable display name for the third e-mail address.
   * https://msdn.microsoft.com/en-us/library/office/cc815669.aspx
   * @readonly
   */
  get email3DisplayName(): string;

  /**
   * Specifies the address type of the third e-mail address.
   * https://msdn.microsoft.com/en-us/library/office/cc842438.aspx
   * @readonly
   */
  get email3AddressType(): string;

  /**
   * Specifies the third e-mail address of the contact.
   * https://msdn.microsoft.com/en-us/library/office/cc815504.aspx
   * @readonly
   */
  get email3EmailAddress(): string;

  /**
   * Specifies the third display name that corresponds to the e-mail address that is specified for the contact.
   * https://msdn.microsoft.com/en-us/library/office/cc815833.aspx
   * @readonly
   */
  get email3OriginalDisplayName(): string;

  /**
   * Specifies the address type for the business fax address for a contact.
   * https://msdn.microsoft.com/en-us/library/office/cc842026.aspx
   * @readonly
   */
  get fax1AddressType(): string;

  /**
   * Specifies the e-mail address of the contact’s business fax.
   * https://msdn.microsoft.com/en-us/library/office/cc765813.aspx
   * @readonly
   */
  get fax1EmailAddress(): string;

  /**
   * Specifies the original display name of the contact’s business fax address.
   * https://msdn.microsoft.com/en-us/library/office/cc765694.aspx
   * @readonly
   */
  get fax1OriginalDisplayName(): string;

  /**
   * Specifies the address type for the contact’s home fax address.
   * https://msdn.microsoft.com/en-us/library/office/cc839741.aspx
   * @readonly
   */
  get fax2AddressType(): string;

  /**
   * Specifies the e-mail address of the contact’s home fax address.
   * https://msdn.microsoft.com/en-us/library/office/cc765668.aspx
   * @readonly
   */
  get fax2EmailAddress(): string;
  /**
   * Specifies the original display name of the contact’s home fax address.
   * https://msdn.microsoft.com/en-us/library/office/cc842101.aspx
   * @readonly
   */
  get fax2OriginalDisplayName(): string;

  /**
   * Specifies the address type for the other contact’s fax address.
   * https://msdn.microsoft.com/en-us/library/office/cc839752.aspx
   * @readonly
   */
  get fax3AddressType(): string;

  /**
   * Specifies the email address of the contact’s other fax address.
   * https://msdn.microsoft.com/en-us/library/office/cc842217.aspx
   * @readonly
   */
  get fax3EmailAddress(): string;

  /**
   * Specifies the original display name of the contact’s other fax address.
   * https://msdn.microsoft.com/en-us/library/office/cc765682.aspx
   * @readonly
   */
  get fax3OriginalDisplayName(): string;

  /**
   * Specifies a URL path from which a client can retrieve free/busy information for the contact as an iCal file, as specified in [MS-OXCICAL].
   * https://msdn.microsoft.com/en-us/library/office/cc765766.aspx
   * @readonly
   */
  get freeBusyLocation(): string;

  /**
   * Contains the birthday of the contact.
   * https://msdn.microsoft.com/en-us/library/office/cc842301.aspx
   * @readonly
   */
  get birthday(): Date | null;

  /**
   * Contains the date of a user's wedding anniversary.
   * https://msdn.microsoft.com/en-us/library/office/cc842132.aspx
   * @readonly
   */
  get anniversary(): Date | null;

  /**
   * Specifies the phonetic pronunciation of the surname of the contact.
   */
  get yomiLastName(): string;

  /**
   * Specifies the phonetic pronunciation of the contact's given name.
   */
  get yomiFirstName(): string;

  /**
   * Specifies the phonetic pronunciation of the contact's company name.
   */
  get yomiCompanyName(): string;

  /**
   * JSON stringify the object properties.
   */
  toJSON(): any;
}