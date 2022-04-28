declare interface ISharePointOnlineQuickAssistWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  WebPartTitle: string;   
  SearchIssue:string;
  SpecifiedDocument:string;
  SpecifiedSite:string;
  UserProfileIssues:string;
  Photosync:string;
  JobTitlesync:string;
  Userinformationsync:string;
  OneDriveIssues:string;
  OneDrivelockicon:string;
  ListLibraryIssues:string;
  MissingForms:string;
  Uneditablewikipage:string;
  Site:string;
  RestoreItems:string;
  GetFileChanges:string;
  Permissionissue:string;
  SelectIssueTip:string;
  AffectedUser:string;
  AffectedSiteLoadList:string;
  SelectList:string;
  AffectedDocument:string;
  CheckIssues:string;
  ShowRemedySteps:string;
  FailedLoadSiteList:string;
  PleaseSelectList:string;
  Checking:string;

  // strings for permssion check module 
  PC_PermissionUrl:string;
  PC_DocumentsWithoutCheckin:string;
  PC_NoDocumentsWithoutCheckin:string;
  PC_ApproveStatusIs:string;
  PC_FileExistingMsg:string;
  PC_FileNotExistingMsg:string;
  PC_PageCustomized:string;
  PC_UserHasPermssionOnDocument:string;
  PC_UserHasNoPermssionOnDocument:string;
  PC_DocumentIsInDraft:string;
  PC_DocumentIsNotInDraft:string;
  PC_ListSecurityLevelHasIssue:string;
  PC_ListSecurityLevelHasNoIssue:string;
  PC_LockDownEnabled:string;
  PC_LockDownNotEnabled:string;
  PC_HasViewPermissionOnList:string;
  PC_HasNoViewPermissionOnList:string;
}

declare module 'SharePointOnlineQuickAssistWebPartStrings' {
  const strings: ISharePointOnlineQuickAssistWebPartStrings;
  export = strings;
}
