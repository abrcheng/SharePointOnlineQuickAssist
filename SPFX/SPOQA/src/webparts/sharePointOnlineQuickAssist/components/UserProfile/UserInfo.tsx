import * as React from 'react';
import {  
    PrimaryButton,    
    TextField,
    Label
  } from 'office-ui-fabric-react/lib/index';
import GraphAPIHelper from '../../../Helpers/GraphAPIHelper';  
import RestAPIHelper from '../../../Helpers/RestAPIHelper';
import SPOQASpinner from '../../../Helpers/SPOQASpinner';
import SPOQAHelper from '../../../Helpers/SPOQAHelper';
import { ISharePointOnlineQuickAssistProps } from '../ISharePointOnlineQuickAssistProps';
import styles from '../SharePointOnlineQuickAssist.module.scss';
import * as strings from 'SharePointOnlineQuickAssistWebPartStrings';

export default class UserInfoQA extends React.Component<ISharePointOnlineQuickAssistProps>
{
  public state = {
    email: this.props.currentUser.loginName,
    isGroup:false,
    isNeedFix:false,
    isChecked:false,
    userId:-1,
    affectedSite:this.props.webAbsoluteUrl    
  };
  private resRef= React.createRef<HTMLDivElement>();  
  private groupInfo1=null;
  private userInfo1=null;
  constructor(props) {
    super(props);
    // this.resRef = React.createRef();
  }

  public render():React.ReactElement<ISharePointOnlineQuickAssistProps>
  {
      return (
        <div>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <TextField
                      label={strings.UI_Label_AffectedSite}
                      multiline={false}
                      onChange={(e)=>{let text:any = e.target; this.setState({affectedSite:text.value});}}
                      value={this.state.affectedSite}
                      required={true}                        
                /> 
              <TextField
                      
                      label={strings.UI_Label_Email}
                      multiline={false}
                      onChange={(e)=>{let text:any = e.target; this.setState({email:text.value});}}
                      value={this.state.email}
                      required={true}                                                
                />   
                <Label>e.g. John@contoso.com </Label>                  
                <div id="UserInfoSyncDiagnoseResult">
                      {this.state.isChecked?<Label>{strings.SS_DiagnoseResultLabel}</Label>:null}
                      <div style={{marginLeft:20}} id="UserInfoSyncDiagnoseResultDiv" ref={this.resRef}>

                      </div>
                </div>
                <PrimaryButton
                    text={strings.UI_CheckIssueforUser}
                    style={{ display: 'inline', marginTop: '10px' }}
                    onClick={() => {this.Check();}}
                  />
                  
                  { this.state.isNeedFix ? 
                  <PrimaryButton
                    text={strings.UI_FixIssues}
                    style={{ display: 'inline', marginTop: '10px', marginLeft:"20px"}}             
                    onClick={() => {this.Fix();}}
                  />: null}
              </div>
          </div>
      </div>
      );
  }

  public async Check()
  {
      // reset status 
      this.ResetStatus();
      var redStyle = "color:red";
      var greenStyle = "color:green";
      if(this.state.affectedSite =="" || !this.state.affectedSite || !SPOQAHelper.ValidateUrl(this.state.affectedSite))
      {
        SPOQAHelper.ShowMessageBar("Error", strings.UI_NonAffectedSite);          
        return;
      }

      if(this.state.email =="" || !this.state.email || !SPOQAHelper.ValidateEmail(this.state.email))
      {
        SPOQAHelper.ShowMessageBar("Error", strings.UI_NonAffectedUser);
        return;
      }      

      SPOQASpinner.Show(strings.SS_Message_Checking);
      console.log(`Start to check sync for eamil ${this.state.email}`);
      try
      {
          // Try to get group via email address
          var groupInfo = await GraphAPIHelper.GetGroupByEmail(this.props.msGraphClient, this.state.email);
          
          // If can't find group with the email address, then try to get user via mail
          if(groupInfo.value.length == 0)
          {
            var userInfo = await GraphAPIHelper.GetUserByEmail(this.props.msGraphClient, this.state.email);
            if(userInfo.value.length ==0) // can't find group and user
            {
              SPOQAHelper.ShowMessageBar("Error", `${strings.UI_NonAADUser} ${this.state.email}!`); 
            }
            else
            {
              this.resRef.current.innerHTML += `<span style='${greenStyle}'>${strings.UI_OneUserAAD} ${userInfo.value[0].userPrincipalName}</span>`;
              // Get user from user info list by userPrincipalName
              var userFromList = await RestAPIHelper.GetUserFromUserInfoList(userInfo.value[0].userPrincipalName, this.props.spHttpClient, this.state.affectedSite);
              if(userFromList == null)
              {
                SPOQAHelper.ShowMessageBar("Error", `${strings.UI_NonUserListUser} ${userInfo.value[0].userPrincipalName}!`);
              }
              else
              {
                // compare the display name, mail, jobtitle, mobilePhone
                this.setState({isChecked:true});
                this.state.userId = userFromList.Id;
                this.userInfo1 = userInfo.value[0];
                var displayNameMismatched = this.userInfo1.displayName != userFromList.Title;
                this.resRef.current.innerHTML+=`</br><span style='${displayNameMismatched? redStyle:greenStyle}'>${strings.UI_Result1} ${this.userInfo1.displayName}, ${strings.UI_Result2} ${userFromList.Title}</span>`;
                var emailMismatched = this.userInfo1.mail != userFromList.EMail;
                this.resRef.current.innerHTML+=`</br><span style='${emailMismatched? redStyle:greenStyle}'>${strings.UI_Result3} ${this.userInfo1.mail}, ${strings.UI_Result4} ${userFromList.EMail}</span>`;
                var jobTitleMismatched = this.userInfo1.jobTitle != userFromList.JobTitle;
                this.resRef.current.innerHTML+=`</br><span style='${jobTitleMismatched? redStyle:greenStyle}'>${strings.UI_Result5} ${this.userInfo1.jobTitle}, ${strings.UI_Result6} ${userFromList.JobTitle}</span>`;
                var workPhoneMismatched = false;
                if(this.userInfo1.businessPhones && this.userInfo1.businessPhones.length && this.userInfo1.businessPhones.length >0)
                {
                  workPhoneMismatched = this.userInfo1.businessPhones[0] != userFromList.WorkPhone;
                  this.userInfo1.WorkPhone = this.userInfo1.businessPhones[0];
                  this.resRef.current.innerHTML+=`</br><span style='${workPhoneMismatched? redStyle:greenStyle}'>${strings.UI_Result7} ${this.userInfo1.WorkPhone}, ${strings.UI_Result8} ${userFromList.WorkPhone}</span>`;
                }
                
                this.setState({isNeedFix:displayNameMismatched||emailMismatched||jobTitleMismatched||workPhoneMismatched});
              }
            }
          }
          else
          {
            this.state.isGroup = true;
            this.groupInfo1 = groupInfo.value[0];
            var groupType = this.groupInfo1.groupTypes.length ==0 ?"null":this.groupInfo1.groupTypes[0];
            this.resRef.current.innerHTML += `<span style='${greenStyle}'>${strings.UI_OneGroupAAD}, securityEnabled=${this.groupInfo1.securityEnabled}, groupType=${groupType}</span>`;
             // Get group from user info list by group Id
             var groupFromList = await RestAPIHelper.GetGroupFromUserInfoList(this.groupInfo1.id, this.props.spHttpClient, this.state.affectedSite);
             if(groupFromList == null)
              {
                SPOQAHelper.ShowMessageBar("Error", `${strings.UI_NonUserListGroup} ${this.groupInfo1.id}!`);
              }
              else
              {
                this.setState({isChecked:true});

                // compare the display name, mail
                this.state.userId = groupFromList.Id;
                if(groupType =="Unified")
                {
                  this.groupInfo1.displayName = this.groupInfo1.displayName + " Members";
                }
                var groupDisplayNameMismatched = this.groupInfo1.displayName != groupFromList.Title;
                this.resRef.current.innerHTML+=`</br><span style='${groupDisplayNameMismatched? redStyle:greenStyle}'>${strings.UI_Result1} ${this.groupInfo1.displayName}, ${strings.UI_Result2} ${groupFromList.Title}</span>`;
                var groupEmailMismatched = this.groupInfo1.mail != groupFromList.EMail;
                this.resRef.current.innerHTML+=`</br><span style='${groupEmailMismatched? redStyle:greenStyle}'>${strings.UI_Result3} ${this.groupInfo1.mail}, ${strings.UI_Result4} ${groupFromList.EMail}</span>`;
                this.setState({isNeedFix:groupEmailMismatched||groupDisplayNameMismatched});
              }
          }
      }
      catch(err)
      {
        console.error(err);
      }

      SPOQASpinner.Hide();
  }  

  public async Fix()
  {
    SPOQASpinner.Show(`${strings.UI_FixEmail} ${this.state.email} ...`);
      var properties: Array<any>=[];
      if(!this.state.isGroup)
      {
        properties.push({key:"WorkPhone", value:this.userInfo1.WorkPhone});
        properties.push({key:"EMail", value:this.userInfo1.mail});
        properties.push({key:"Title", value:this.userInfo1.displayName});
        properties.push({key:"JobTitle", value:this.userInfo1.jobTitle});
      }
      else
      {     
        properties.push({key:"EMail", value:this.groupInfo1.mail});
        properties.push({key:"Title", value:this.groupInfo1.displayName});
      }

      RestAPIHelper.FixUserInfoItem(this.state.userId, this.props.spHttpClient, this.state.affectedSite, properties, this.SuccessCallBack, this.FailedCallback);
  }

  public SuccessCallBack()
  {    
      SPOQASpinner.Hide();
      SPOQAHelper.ShowMessageBar("Success", `${strings.UI_FixSuccess}`); 
  }

  public FailedCallback()
  {
      SPOQASpinner.Hide();
      SPOQAHelper.ShowMessageBar("Error", `${strings.UI_FixFailed}`);     
  }

  private ResetStatus():void
    {      
      this.state.userId = -1; 
      this.resRef.current.innerHTML ="";
      this.state.isGroup = false;
      SPOQAHelper.ResetFormStaus();       
    }  

}