declare interface IMailboxWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  AddressFieldLabel: string;
  AppAddressFieldLabel: string;
  NameFieldLabel: string;
  SpWebBaseUrlFieldLabel: string;
  SpDocLibIdFieldLabel: string;
  SpListIdFieldLabel: string;
  MembersFieldLabel: string;
  BackEndApiFieldLabel: string;
}

declare module 'MailboxWebPartStrings' {
  const strings: IMailboxWebPartStrings;
  export = strings;
}
