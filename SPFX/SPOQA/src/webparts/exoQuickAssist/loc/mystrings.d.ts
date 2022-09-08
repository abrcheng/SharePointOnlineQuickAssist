declare interface IExoQuickAssistWebPartStrings {
   WelcomeToEXOQA:string,
   OrganizationSettings:string,
   PeopleInsights:string,
   SelectIssueTip:string,
   CheckIssue:string,
   ShowRemedySteps:string,
   AffectedUser:string   
}

declare module 'ExoQuickAssistWebPartStrings' {
  const strings: IExoQuickAssistWebPartStrings;
  export = strings;
}
