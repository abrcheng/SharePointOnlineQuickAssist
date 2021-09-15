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
  //PrimaryButton,
  SelectableOptionMenuItemType
} from 'office-ui-fabric-react/lib/index';
import SearchDocumentQA from './SearchDocument';
import SearchLibraryQA from './SearchLibrary';
import SearchSiteQA from './SearchSite';
import UserProfilePhotoQA from './UserProfilePhoto';
import UserProfileEmailQA from './UserProfileEmail';
import UserProfileManagerQA from './UserProfileManager';
import UserProfileTitleQA from './UserProfileTitle';
import UserProfileDepartmentQA from './UserProfileDepartment';
import SearchPeopleQA from './SearchPeople';

//import { Button } from 'office-ui-fabric-react/lib/Button';
// https://developer.microsoft.com/en-us/fluentui?fabricVer=6#/controls/web/combobox
const INITIAL_OPTIONS: IComboBoxOption[] = [
  { key: 'Search', text: 'Search Issues', itemType: SelectableOptionMenuItemType.Header },
  { key: 'SearchDocument', text: 'Specified Document' },
  { key: 'SearchPeople', text: 'People' },
  { key: 'SearchLibrary', text: 'Specified Library' },
  { key: 'SearchSite', text: 'Specified Site' },  
  { key: 'UserProfile', text: 'User Profile Issues', itemType: SelectableOptionMenuItemType.Header },
  { key: 'UserProfilePhoto', text: 'Photo sync issue' },
  { key: 'UserProfileTitle', text: 'Title sync issue'},
  { key: 'UserProfileEmail', text: 'Email sync issue' },
  { key: 'UserProfileManager', text: 'Manager sync issue' },
  { key: 'UserProfileDepartment', text: 'Department sync issue' }  
];

const wrapperClassName = mergeStyles({
  selectors: {
    '& > *': { marginBottom: '20px' },
    '& .ms-ComboBox': { maxWidth: '300px' }
  }
});

export default class SharePointOnlineQuickAssist extends React.Component<ISharePointOnlineQuickAssistProps, {}> {
  public state = {
    selectedKey: ""
  };

  public render(): React.ReactElement<ISharePointOnlineQuickAssistProps> {

    const sPOQADetail = () => {
      switch(this.state.selectedKey) {
        case "SearchDocument":   return <SearchDocumentQA />;
        case "SearchPeople":   return <SearchPeopleQA />; 
        case "SearchLibrary":   return <SearchLibraryQA />;
        case "SearchSite":   return <SearchSiteQA />;       
        case "UserProfilePhoto":   return <UserProfilePhotoQA />; 
        case "UserProfileTitle":   return <UserProfileTitleQA />;
        case "UserProfileEmail":   return <UserProfileEmailQA />; 
        case "UserProfileManager":   return <UserProfileManagerQA />;
        case "UserProfileDepartment":   return <UserProfileDepartmentQA />; 
        default: return <div id="NoContentPlaceHolder"/>;
      }
    };

    return (
      <div className={ styles.sharePointOnlineQuickAssist }>
         <Fabric className={wrapperClassName}>
        <div className={ styles.container }>
          <div className={ styles.row } id="SPOQAHeaderContainer">
            <div className={ styles.column }>
              <span className={ styles.title }>Welcome to use SharePoint Online Quick Assist!</span>                            
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
        </div>
        </Fabric>
      </div>
    );
  }
}
