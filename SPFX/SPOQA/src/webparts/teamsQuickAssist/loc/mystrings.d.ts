declare interface ITeamsQuickAssistWebPartStrings {
  Admin: string;
  Test: string;  
  WebPartTitle:string;
  SelectIssueTip:string;
  RO_CheckPass:string;
  RO_CheckFail:string;
  RO_APIError:string;
  RO_Remedy:string;
  AccountIssue:string;
  LoginCookieError:string;
}

declare module 'TeamsQuickAssistWebPartStrings' {
  const strings: ITeamsQuickAssistWebPartStrings;
  export = strings;
}
