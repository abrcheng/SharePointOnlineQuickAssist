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
export default class UserInfoQA extends React.Component<ISharePointOnlineQuickAssistProps>
{
  public state = {
    email: "",
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
          <TextField
                  label="Affected Site:"
                  multiline={false}
                  onChange={(e)=>{let text:any = e.target; this.setState({affectedSite:text.value});}}
                  value={this.state.affectedSite}
                  required={true}                        
            /> 
          <TextField
                  label="Email of Affected User/Group:"
                  multiline={false}
                  onChange={(e)=>{let text:any = e.target; this.setState({email:text.value});}}
                  value={this.state.email}
                  required={true}                                                
            />                  
            <div id="UserInfoSyncDiagnoseResult">
                  {this.state.isChecked?<Label>Diagnose result:</Label>:null}
                  <div id="UserInfoSyncDiagnoseResultDiv" ref={this.resRef}>

                  </div>
            </div>
            <PrimaryButton
                text="Check Sync Issue"
                style={{ display: 'inline', marginTop: '10px' }}
                onClick={() => {this.Check();}}
              />
              
              { this.state.isNeedFix ? 
              <PrimaryButton
                text="Fix It"
                style={{ display: 'inline', marginTop: '10px', marginLeft:"20px"}}             
                onClick={() => {this.Fix();}}
              />: null}
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
        SPOQAHelper.ShowMessageBar("Error", "Affected site can't be null or invalid!");          
        return;
      }

      if(this.state.email =="" || !this.state.email || !SPOQAHelper.ValidateEmail(this.state.email))
      {
        SPOQAHelper.ShowMessageBar("Error", "Affected user can't be null or invalid!");
        return;
      }      

      SPOQASpinner.Show("Checking ...");
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
              SPOQAHelper.ShowMessageBar("Error", `In the AAD can't find any group or user with the email address ${this.state.email}!`); 
            }
            else
            {
              this.resRef.current.innerHTML += `<span style='${greenStyle}'>Get one user ${userInfo.value[0].userPrincipalName} from AAD.</span>`;
              // Get user from user info list by userPrincipalName
              var userFromList = await RestAPIHelper.GetUserFromUserInfoList(userInfo.value[0].userPrincipalName, this.props.spHttpClient, this.state.affectedSite);
              if(userFromList == null)
              {
                SPOQAHelper.ShowMessageBar("Error", `In the user info list can't find any user with the UPN ${userInfo.value[0].userPrincipalName}!`);
              }
              else
              {
                // compare the display name, mail, jobtitle, mobilePhone
                this.setState({isChecked:true});
                this.state.userId = userFromList.Id;
                this.userInfo1 = userInfo.value[0];
                var displayNameMismatched = this.userInfo1.displayName != userFromList.Title;
                this.resRef.current.innerHTML+=`</br><span style='${displayNameMismatched? redStyle:greenStyle}'>Display name in AAD is ${this.userInfo1.displayName}, display name in user info list is ${userFromList.Title}</span>`;
                var emailMismatched = this.userInfo1.mail != userFromList.EMail;
                this.resRef.current.innerHTML+=`</br><span style='${emailMismatched? redStyle:greenStyle}'>Email in AAD is ${this.userInfo1.mail}, email in user info list is ${userFromList.EMail}</span>`;
                var jobTitleMismatched = this.userInfo1.jobTitle != userFromList.JobTitle;
                this.resRef.current.innerHTML+=`</br><span style='${jobTitleMismatched? redStyle:greenStyle}'>JobTitle in AAD is ${this.userInfo1.jobTitle}, JobTitle in user info list is ${userFromList.JobTitle}</span>`;
                var workPhoneMismatched = false;
                if(this.userInfo1.businessPhones && this.userInfo1.businessPhones.length && this.userInfo1.businessPhones.length >0)
                {
                  workPhoneMismatched = this.userInfo1.businessPhones[0] != userFromList.WorkPhone;
                  this.userInfo1.WorkPhone = this.userInfo1.businessPhones[0];
                  this.resRef.current.innerHTML+=`</br><span style='${workPhoneMismatched? redStyle:greenStyle}'>WorkPhone in AAD is ${this.userInfo1.WorkPhone}, WorkPhone in user info list is ${userFromList.WorkPhone}</span>`;
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
            this.resRef.current.innerHTML += `<span style='${greenStyle}'>Get one group from AAD, securityEnabled=${this.groupInfo1.securityEnabled}, groupType=${groupType}</span>`;
             // Get group from user info list by group Id
             var groupFromList = await RestAPIHelper.GetGroupFromUserInfoList(this.groupInfo1.id, this.props.spHttpClient, this.state.affectedSite);
             if(groupFromList == null)
              {
                SPOQAHelper.ShowMessageBar("Error", `In the user info list can't find any group with group Id ${this.groupInfo1.id}!`);
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
                this.resRef.current.innerHTML+=`</br><span style='${groupDisplayNameMismatched? redStyle:greenStyle}'>Display name in AAD is ${this.groupInfo1.displayName}, display name in user info list is ${groupFromList.Title}</span>`;
                var groupEmailMismatched = this.groupInfo1.mail != groupFromList.EMail;
                this.resRef.current.innerHTML+=`</br><span style='${groupEmailMismatched? redStyle:greenStyle}'>Email in AAD is ${this.groupInfo1.mail}, email in user info list is ${groupFromList.EMail}</span>`;
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
    SPOQASpinner.Show(`Fixing user info item with eamil ${this.state.email} in the site ...`);
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
      SPOQAHelper.ShowMessageBar("Success", "Fix in the user info list completed, please recheck for verfiying it."); 
  }

  public FailedCallback()
  {
      SPOQASpinner.Hide();
      SPOQAHelper.ShowMessageBar("Error", "Fix in the user info list failed.");     
  }

  private ResetStatus():void
    {      
      this.state.userId = -1; 
      this.resRef.current.innerHTML ="";
      this.state.isGroup = false;
      SPOQAHelper.ResetFormStaus();       
    }  

}