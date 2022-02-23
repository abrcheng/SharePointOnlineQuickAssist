import * as React from 'react';
import styles from './SharePointOnlineQuickAssist.module.scss';
import { ISharePointOnlineQuickAssistProps } from './ISharePointOnlineQuickAssistProps';
import { escape } from '@microsoft/sp-lodash-subset';
//import '~@microsoft/sp-office-ui-fabric-core/dist/sass/SPFabricCore.scss';
import {
  ComboBox,
  Fabric,
  IComboBox,
  IComboBoxOption,
  mergeStyles, 
  SelectableOptionMenuItemType,
  Spinner,
  MessageBar,
  MessageBarType
} from 'office-ui-fabric-react/lib/index';
import SearchDocumentQA from './Search/SearchDocument';
import SearchLibraryQA from './Search/SearchLibrary';
import SearchSiteQA from './Search/SearchSite';
import UserProfilePhotoQA from './UserProfile/UserProfilePhoto';
import UserProfileEmailQA from './UserProfile/UserProfileEmail';
import UserProfileManagerQA from './UserProfile/UserProfileManager';
import UserProfileTitleQA from './UserProfile/UserProfileTitle';
import UserProfileDepartmentQA from './UserProfile/UserProfileDepartment';
import SearchPeopleQA from './Search/SearchPeople';
import UserInfoQA from './UserProfile/UserInfo'; 
import RestoreItemsQA from './Site/RestoreItems';
import OneDriveLockIconQA from './OneDrive/OneDriveLockIcon';
import RepairFormQA from './List/RepairListForms';
import { WebPartContext } from "@microsoft/sp-webpart-base"; 
import { SPComponentLoader } from '@microsoft/sp-loader';
import SPOQAHelper from '../../Helpers/SPOQAHelper';
import * as strings from 'SharePointOnlineQuickAssistWebPartStrings';

//import { Button } from 'office-ui-fabric-react/lib/Button';
// https://developer.microsoft.com/en-us/fluentui?fabricVer=6#/controls/web/combobox
const INITIAL_OPTIONS: IComboBoxOption[] = [
  { key: 'Search', text: 'Search Issues', itemType: SelectableOptionMenuItemType.Header },
  { key: 'SearchDocument', text: '  Specified Document' },
  // { key: 'SearchPeople', text: 'People' },
  //{ key: 'SearchLibrary', text: 'Specified Library' },
  { key: 'SearchSite', text: '  Specified Site' },  
  { key: 'UserProfile', text: 'User Profile Issues', itemType: SelectableOptionMenuItemType.Header },
  { key: 'UserProfilePhoto', text: '  Photo sync' },  
  { key: 'UserProfileTitle', text: '  Job Title sync'},
  { key: 'UserInfoSync', text: '  User information sync'},
  // { key: 'UserProfileEmail', text: 'Email sync issue' },
  // { key: 'UserProfileManager', text: 'Manager sync issue' },
  // { key: 'UserProfileDepartment', text: 'Department sync issue' }  
  { key: 'OneDrive', text: 'OneDrive Issues', itemType: SelectableOptionMenuItemType.Header },
  { key: 'OneDriveLockIcon', text: 'OneDrive lock icon' }, 
  { key: 'List', text: 'List/Library Issues', itemType: SelectableOptionMenuItemType.Header },
  { key: 'ListMissingForm', text: 'Missing New/Disp/Edit Forms' }, 
  { key: 'Site', text: 'Site', itemType: SelectableOptionMenuItemType.Header },
  { key: 'Restore', text: 'Restore' }, 
];

const wrapperClassName = mergeStyles({
  selectors: {
    '& > *': { marginBottom: '20px' },
    '& .ms-ComboBox': { maxWidth: '300px' },
    '& .ms-ComboBox-option':{marginLeft:"15px"}
  }
});

export default class SharePointOnlineQuickAssist extends React.Component<ISharePointOnlineQuickAssistProps, {}> {
  public state = {
    selectedKey: ""
  };
  
  public componentDidMount(): void
  {
    SPComponentLoader.loadScript('/_layouts/15/init.js', {
      globalExportsName: '$_global_init'
    })
    .then((): Promise<{}> => {
      return SPComponentLoader.loadScript('/_layouts/15/MicrosoftAjax.js', {
        globalExportsName: 'Sys'
      });
    })
    .then((): Promise<{}> => {
      return SPComponentLoader.loadScript('/_layouts/15/SP.Runtime.js', {
        globalExportsName: 'SP'
      });
    })
    .then((): Promise<{}> => {
      return SPComponentLoader.loadScript('/_layouts/15/SP.js', {
        globalExportsName: 'SP'
      });
    });
  }

  public render(): React.ReactElement<ISharePointOnlineQuickAssistProps> {
    // this.props.webPartContext
    const sPOQADetail = () => {
      switch(this.state.selectedKey) {
        case "SearchDocument":   return <SearchDocumentQA spHttpClient={this.props.spHttpClient} msGraphClient={this.props.msGraphClient} webUrl={this.props.webUrl} webAbsoluteUrl={this.props.webAbsoluteUrl} rootUrl={this.props.rootUrl} currentUser={this.props.currentUser} ctx={this.props.ctx}/>;
        case "SearchPeople":   return <SearchPeopleQA spHttpClient={this.props.spHttpClient} msGraphClient={this.props.msGraphClient} webUrl={this.props.webUrl} webAbsoluteUrl={this.props.webAbsoluteUrl} rootUrl={this.props.rootUrl} currentUser={this.props.currentUser} ctx={this.props.ctx}/>; 
        case "SearchLibrary":   return <SearchLibraryQA spHttpClient={this.props.spHttpClient} msGraphClient={this.props.msGraphClient} webUrl={this.props.webUrl} webAbsoluteUrl={this.props.webAbsoluteUrl} rootUrl={this.props.rootUrl} currentUser={this.props.currentUser} ctx={this.props.ctx}/>;
        case "SearchSite":   return <SearchSiteQA spHttpClient={this.props.spHttpClient} msGraphClient={this.props.msGraphClient} webUrl={this.props.webUrl} webAbsoluteUrl={this.props.webAbsoluteUrl} rootUrl={this.props.rootUrl} currentUser={this.props.currentUser} ctx={this.props.ctx}/>;       
        case "UserProfilePhoto":   return <UserProfilePhotoQA spHttpClient={this.props.spHttpClient} msGraphClient={this.props.msGraphClient} webUrl={this.props.webUrl} webAbsoluteUrl={this.props.webAbsoluteUrl} rootUrl={this.props.rootUrl} currentUser={this.props.currentUser} ctx={this.props.ctx}/>; 
        case "UserProfileTitle":   return <UserProfileTitleQA spHttpClient={this.props.spHttpClient} msGraphClient={this.props.msGraphClient} webUrl={this.props.webUrl} webAbsoluteUrl={this.props.webAbsoluteUrl} rootUrl={this.props.rootUrl} currentUser={this.props.currentUser} ctx={this.props.ctx}/>;
        case "UserProfileEmail":   return <UserProfileEmailQA spHttpClient={this.props.spHttpClient} msGraphClient={this.props.msGraphClient} webUrl={this.props.webUrl} webAbsoluteUrl={this.props.webAbsoluteUrl} rootUrl={this.props.rootUrl} currentUser={this.props.currentUser} ctx={this.props.ctx}/>; 
        case "UserProfileManager":   return <UserProfileManagerQA spHttpClient={this.props.spHttpClient} msGraphClient={this.props.msGraphClient} webUrl={this.props.webUrl} webAbsoluteUrl={this.props.webAbsoluteUrl} rootUrl={this.props.rootUrl} currentUser={this.props.currentUser} ctx={this.props.ctx}/>;
        case "UserProfileDepartment":   return <UserProfileDepartmentQA spHttpClient={this.props.spHttpClient} msGraphClient={this.props.msGraphClient} webUrl={this.props.webUrl} webAbsoluteUrl={this.props.webAbsoluteUrl} rootUrl={this.props.rootUrl} currentUser={this.props.currentUser} ctx={this.props.ctx}/>; 
        case "UserInfoSync":   return <UserInfoQA spHttpClient={this.props.spHttpClient} msGraphClient={this.props.msGraphClient} webUrl={this.props.webUrl} webAbsoluteUrl={this.props.webAbsoluteUrl} rootUrl={this.props.rootUrl} currentUser={this.props.currentUser} ctx={this.props.ctx}/>; 
        case "OneDriveLockIcon":   return <OneDriveLockIconQA spHttpClient={this.props.spHttpClient} msGraphClient={this.props.msGraphClient} webUrl={this.props.webUrl} webAbsoluteUrl={this.props.webAbsoluteUrl} rootUrl={this.props.rootUrl} currentUser={this.props.currentUser} ctx={this.props.ctx}/>; 
        case "ListMissingForm":   return <RepairFormQA spHttpClient={this.props.spHttpClient} msGraphClient={this.props.msGraphClient} webUrl={this.props.webUrl} webAbsoluteUrl={this.props.webAbsoluteUrl} rootUrl={this.props.rootUrl} currentUser={this.props.currentUser} ctx={this.props.ctx}/>; 
        case "Restore": return <RestoreItemsQA spHttpClient={this.props.spHttpClient} msGraphClient={this.props.msGraphClient} webUrl={this.props.webUrl} webAbsoluteUrl={this.props.webAbsoluteUrl} rootUrl={this.props.rootUrl} currentUser={this.props.currentUser} ctx={this.props.ctx}/>; 
        default: return <div id="NoContentPlaceHolder"/>;
      }
    };

    return (
      <div className={ styles.sharePointOnlineQuickAssist }>
         <Fabric className={wrapperClassName} id="SPOAQFabric">
        <div className={ styles.container }>
          <div className={ styles.row } id="SPOQAHeaderContainer">
            <div className={ styles.column }>
              <span className={ styles.title }>{strings.WebPartTitle}</span>                            
            </div>
          </div>

          <div className={ styles.row } id="SPOQAQuestionsContainer">
            <div className={ styles.column }>                         
                  <div>                    
                    <ComboBox
                      defaultSelectedKey="-1"
                      label="Please select issue which you want to check"
                      allowFreeform
                      autoComplete="on"                      
                      options={INITIAL_OPTIONS} 
                      required={true}                    
                      onChange ={(ev: React.FormEvent<IComboBox>, option?: IComboBoxOption): void => {
                        this.setState({ selectedKey: option.key});}} 
                    />                  
                  </div>                 
            </div>
          </div>
          
          <div className={ styles.row } id="SPOQADetailContainer">
            <div className={ styles.column }>
              {sPOQADetail()}
            </div>
          </div>
          <div className={ styles.row } id="SPOQAStatusContainer">
            <div className={ styles.column }>
              <div>        
                <Spinner id="SPOQASpinner" label="Checking..." ariaLive="assertive" labelPosition="left" style={{display:"none"}} />
                <div id="SPOQAErrorMessageBarContainer" style={{display:"none"}}>
                  <MessageBar id="SPOQAErrorMessageBar" messageBarType={MessageBarType.error} isMultiline={true} onDismiss={()=>{SPOQAHelper.Hide("SPOQAErrorMessageBarContainer");}} dismissButtonAriaLabel="Close" >
                              SPOQAErrorMessageBar
                  </MessageBar>
                </div>
                <div id="SPOQASuccessMessageBarContainer" style={{display:"none"}}>
                  <MessageBar id="SPOQASuccessMessageBar" messageBarType={MessageBarType.success} isMultiline={true} onDismiss={()=>{SPOQAHelper.Hide("SPOQASuccessMessageBarContainer");}} dismissButtonAriaLabel="Close" >
                          SPOQASuccessMessageBar
                  </MessageBar>
                </div>
                <div id="SPOQAWarningMessageBarContainer" style={{display:"none"}}>
                  <MessageBar id="SPOQAWarningMessageBar" messageBarType={MessageBarType.warning} isMultiline={true} onDismiss={()=>{SPOQAHelper.Hide("SPOQAWarningMessageBarContainer");}} dismissButtonAriaLabel="Close">
                          SPOQAWarningMessageBar
                  </MessageBar>
                </div>
                <div id="SPOQAInfoMessageBarContainer" style={{display:"none"}}>
                  <MessageBar id="SPOQAInfoMessageBar" messageBarType={MessageBarType.info} isMultiline={true} onDismiss={()=>{SPOQAHelper.Hide("SPOQAInfoMessageBarContainer");}} dismissButtonAriaLabel="Close">
                        SPOQAInfoMessageBar
                  </MessageBar>
                </div>
              </div>
            </div>
          </div>
        </div>
        </Fabric>
      </div>
    );
  }
}
