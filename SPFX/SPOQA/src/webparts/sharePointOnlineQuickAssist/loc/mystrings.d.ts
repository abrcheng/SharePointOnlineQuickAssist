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
  RemedySteps:string;
  CanBeFixedIn:string;
  ThisPage:string;
  AffectedSite:string;

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
  PC_TryDisable3PCode:string;
  PC_LackPermissionOn:string;
  PC_CanNotLoad:string;

  // strings for Restore Items
  RI_DeletedBy:string;
  RI_PathFilter:string;
  RI_StartDate:string;
  RI_SelectADate:string;
  RI_EndDate:string;
  RI_QueryItems:string;
  RI_Restore:string;
  RI_Export:string;
  RI_Querying:string;
  RI_Queried:string;
  RI_Items:string;
  RI_Filtered:string;
  RI_In:string;
  RI_RestoreItemFrom:string;
  RI_Seconds:string;
  RI_To:string;
  RI_PleaseWait:string;
  RI_Restored:string;
  RI_WithErrorMessage:string;
  RI_DeletedDate:string;
  RI_Path:string;
  RI_DeletedByEmail:string;
  RI_DeletedByName:string;
  RI_NoData:string;
  RI_FullDocumentPath:string;
  RI_NotMatchListRootFolder:string;
  RI_FailedToLoadItem:string;
  RI_URLInvalid:string;
  RI_ItemTypeUnKnow:string;
  RI_FailedToGetData:string;

  // Strings for Search document
  SD_DocumentPathCanNotBeNull:string;
  SD_DocumentCanBeSearched:string;
  SD_SearchByFullPathException:string;
  SD_IsWebNoCrawlException:string;
  SD_DectectedNocrawlList:string;
  SD_IsListNoCrawlException:string;
  SD_DectectedDisplayFormIsMissing:string;
  SD_IsListMissDisplayFormException:string;
  SD_TheNocrawlEnabledList:string;
  SD_TheNocrawlNotEnabledList:string;
  SD_TheDisplayFormMissed:string;
  SD_TheDisplayFormNotMissed:string;
  SD_IsDocumentInDraftVersionException:string;
  SD_FolderSkipDraftCheck:string;
  SD_DectectedNocrawlSite:string;
  SD_NocrawlEnabledSite:string;
  SD_NocrawlNotEnabledSite:string;

  // Strings for OneDrive Lock Icon issue
  OL_CheckingIssueForLibrary:string;
  OL_LockdownEanbled:string;
  OL_LockdownNotEanbled:string;
  OL_HasEditPermssionOnLibrary:string;
  OL_LackEditPermssionOnLibrary:string;
  OL_OfflineAvailability:string;
  OL_RequireCheckOut:string;
  OL_DraftItemSecurity:string;
  OL_AnyUserCanRead:string;
  OL_EditUserCanRead:string;
  OL_ApproverCanRead:string;
  OL_ContentApproval:string;
  OL_ValidationFormula:string;
  OL_OfflineAvailabilityForWeb:string;
  OL_SchemaCheckPassed:string;
  OL_ColumnHasBeenSetTo:string;
  OL_Required:string;
  OL_Formula:string;
  OL_ValidationMessage:string;

}

declare module 'SharePointOnlineQuickAssistWebPartStrings' {
  const strings: ISharePointOnlineQuickAssistWebPartStrings;
  export = strings;
}
