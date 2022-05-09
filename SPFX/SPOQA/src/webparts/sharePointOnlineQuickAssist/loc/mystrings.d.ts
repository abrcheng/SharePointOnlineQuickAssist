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

  // Strings for Search Site
  SS_DiagnoseResultLabel:string;
  SS_FoundSite:string;
  SS_SiteNoExist1:string;
  SS_SiteNoExist2:string;
  SS_NoCrawlEnabled1:string;
  SS_NoCrawlEnabled2:string;
  SS_SiteIndexEnabled1:string;
  SS_SiteIndexEnabled2:string;
  SS_HaveAccess:string;
  SS_NoAccess:string;
  SS_InMembers:string;
  SS_NotInMembers:string;
  SS_Label_AffectedSite:string;
  SS_Label_CheckIssues:string;
  SS_Label_FixIssues:string;
  SS_Message_Checking:string;
  SS_Message_NoSearchResult:string;
  SS_Ex_GetWebError:string;
  SS_Ex_IsWebNoCrawlError:string;
  SS_Ex_GetUserInfoError:string;
  SS_Message_ResultDuplicate:string;
  SS_Message_SiteSearchable:string;
  SS_Message_FixSite:string;
  SS_Ex_FixWebNoCrawlError:string;
  SS_Ex_AddUserInMembersError:string;
  SS_Message_FxiedAll:string;
  SS_Message_SearchAndOffline:string;
  SS_Message_WaitAfterFix:string;
  SS_Message_CheckPermissions:string;
  SS_Message_AddInMembers:string;

  // Strings for UserInfo
  UI_Label_AffectedSite:string;
  UI_Label_Email:string;
  UI_CheckIssueforUser:string;
  UI_FixIssues:string;
  UI_Result1:string;
  UI_Result2:string;
  UI_Result3:string;
  UI_Result4:string;
  UI_Result5:string;
  UI_Result6:string;
  UI_Result7:string;
  UI_Result8:string;
  UI_NonAffectedSite:string;
  UI_NonAffectedUser:string;
  UI_FixEmail:string;
  UI_NonAADUser:string;
  UI_NonUserListUser:string;
  UI_OneUserAAD:string;
  UI_OneGroupAAD:string;
  UI_NonUserListGroup:string;
  UI_FixSuccess:string;
  UI_FixFailed:string;

  //Strings for User Profile Department
  UPD_CheckDepartment:string;
  
  //Strings for User Profile Email
  UPE_CheckEmail:string;

  //String for User Profile Manager
  UPM_CheckManager:string;
  
  //Strings for User Profile Photo
  UPP_PhotoURL:string;
  UPP_PhotoAAD:string;
  UPP_PhotoUserProfile:string;
  UPP_PhotoSuccess:string;
  UPP_PhotoFailed:string;
  UPP_NonAADPhoto:string;

  //Strings for User Profile Title
  UPT_AADTitle:string;
  UPT_UserProfileTitle:string;
  UPT_UserInfoListTitle:string;
  UPT_NonAADTitle:string;
  UPT_NonSiteTitle:string;
  UPT_FixSiteTitle:string;
  UPT_FixUserProfileTitle:string;
  UPT_FailedUserProfileTitle:string;
  UPT_SuccessUserInfoListTitle:string;
  UPT_FailedUserInfoListTitle:string;

  // Strings for Missing Forms
  MF_Label_AffectedSite:string;
  MF_Label_SelectList:string;
  MF_Label_DiagnoseResult:string;
  MF_Message_DispFormMiss:string;
  MF_Message_DispFormExist:string;
  MF_Message_NewFormMiss:string;
  MF_Message_NewFormExist:string;
  MF_Message_EditFormMiss:string;
  MF_Message_EditFormExist:string;
  MF_Ex_LoadListsError:string;
  MF_Ex_ListNotSelected:string;
  MF_Message_CheckingForms:string;
  MF_Ex_CheckFormsError:string;
  MF_Message_FixForms:string;
  MF_Ex_FixDispFormError:string;
  MF_Ex_FixNewFormError:string;
  MF_Ex_FixEditFormError:string;
  MF_Message_FixedAll:string;

  // Strings for Get Files Changes
  FC_Label_QuerySite:string;
  FC_Lable_ModifedUser:string;
  FC_Label_PathFilter:string;
  FC_Label_StartDate:string;
  FC_Message_SelectDate:string;
  FC_Label_EndDate:string;
  FC_Label_GetFiles:string;
  FC_Label_Export:string;
  FC_Message_Quering:string;
  FC_Ex_GetFilesChangeError:string;
  FC_Message_QueryDone:string;
  FC_Ex_GetSiteError:string;
}

declare module 'SharePointOnlineQuickAssistWebPartStrings' {
  const strings: ISharePointOnlineQuickAssistWebPartStrings;
  export = strings;
}
