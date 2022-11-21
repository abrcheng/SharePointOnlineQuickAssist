import * as React from 'react';
import {  
    DefaultButton,    
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
export default class UserProfileTitleQA extends React.Component<ISharePointOnlineQuickAssistProps>
{
    public state = {
        affectedUser: this.props.currentUser.loginName,
        aadJobTitle:"",
        uapJobtitle:"",
        siteJobTitle:"",
        userId:-1,
        affectedSite:this.props.webAbsoluteUrl
      };

    public render():React.ReactElement<ISharePointOnlineQuickAssistProps>
    {       
        return (
            <div>
                <div className={ styles.row }>
                    <div className={ styles.column }>
                      <TextField
                              label={strings.AffectedSite}
                              multiline={false}
                              onChange={(e)=>{let text:any = e.target; this.setState({affectedSite:text.value});}}
                              value={this.state.affectedSite}
                              required={true}                        
                        /> 
                      <TextField
                              label={strings.AffectedUser}
                              multiline={false}
                              onChange={(e)=>{let text:any = e.target; this.setState({affectedUser:text.value});}}
                              value={this.state.affectedUser}
                              required={true}         
                                                           
                        />   
                        <Label>e.g. John@contoso.com </Label>                         
                        {this.state.aadJobTitle != ""? <Label>{strings.UPT_AADTitle} <span style={{"color":"Green"}}>{this.state.aadJobTitle}</span></Label> : null}
                        {this.state.aadJobTitle != "" && this.state.userId && this.state.userId !=-1? <Label>{strings.UPT_UserProfileTitle} <span style={this.state.uapJobtitle != this.state.aadJobTitle? {"color":"Red"}:{"color":"Green"}}>{this.state.uapJobtitle}</span></Label>: null}
                        {this.state.aadJobTitle != "" && this.state.userId && this.state.userId !=-1?<Label>{strings.UPT_UserInfoListTitle} <span style={this.state.siteJobTitle != this.state.aadJobTitle? {"color":"Red"}:{"color":"Green"}}>{this.state.siteJobTitle}</span></Label>: null}

                        <DefaultButton
                            text={strings.CheckIssues}
                            style={{ display: 'inline', marginTop: '10px' }}
                            onClick={() => {this.CheckUserProfileTitle();}}
                          />
                          
                          { (this.state.siteJobTitle != this.state.aadJobTitle || this.state.aadJobTitle != this.state.uapJobtitle) && this.state.userId && this.state.userId !=-1 ? 
                          <DefaultButton
                            text={strings.UI_FixIssues}
                            style={{ display: 'inline', marginTop: '10px', marginLeft:"20px"}}
                            hidden={this.state.siteJobTitle == this.state.aadJobTitle && this.state.uapJobtitle == this.state.aadJobTitle}
                            onClick={() => {this.FixJobTitle();}}
                          />: null}
                  </div>  
              </div>
            </div>
        );
    }

    public async CheckUserProfileTitle()
    {
        // reset status 
        this.ResetStatus();
        
        if(this.state.affectedSite =="" || this.state.affectedUser =="" )
        {
          SPOQAHelper.ShowMessageBar("Error", strings.UI_NonAffectedSiteandUser);          
          return;
        }

        if(this.state.affectedSite =="" || !this.state.affectedSite || !SPOQAHelper.ValidateUrl(this.state.affectedSite))
        {
          SPOQAHelper.ShowMessageBar("Error", strings.UI_NonAffectedSite);          
          return;
        }

        if(this.state.affectedUser =="" || !this.state.affectedUser || !SPOQAHelper.ValidateEmail(this.state.affectedUser))
        {
          SPOQAHelper.ShowMessageBar("Error", strings.UI_NonAffectedUser);
          return;
        }      

        SPOQASpinner.Show(strings.Checking);
        console.log("Start to CheckUserProfileTitle");
        try
        {
          var userInfoAAD = await GraphAPIHelper.GetUserInfo(this.state.affectedUser, this.props.msGraphClient);        
          console.log(`Job title from AAD is ${userInfoAAD.jobTitle}`);        
          this.setState({aadJobTitle:userInfoAAD.jobTitle});        
          var userInfoUAP = await RestAPIHelper.GetUserInfoFromUserProfile(this.state.affectedUser,this.props.spHttpClient, this.state.affectedSite);
          console.log(`Job title from UAP is ${userInfoUAP.Title}`);
          this.setState({uapJobtitle:userInfoUAP.Title});
          var userInfoSite = await RestAPIHelper.GetUserFromUserInfoList(this.state.affectedUser, this.props.spHttpClient, this.state.affectedSite);
          console.log(`Job title from user info list is ${userInfoSite.JobTitle}`);
          this.setState({siteJobTitle:userInfoSite.JobTitle});    
          this.setState({userId:userInfoSite.Id});    
          console.log("ended CheckUserProfileTitle");
        }
        catch(err)
        {
           this.forceUpdate();
           if(this.state.aadJobTitle == "" || !this.state.aadJobTitle)
           {
              SPOQAHelper.ShowMessageBar("Error", strings.UPT_NonAADTitle);              
           }
           else if(this.state.userId ==-1)
           {
              SPOQAHelper.ShowMessageBar("Error", strings.UPT_NonSiteTitle);
           }
        }

        SPOQASpinner.Hide();
    }    

    public async FixJobTitle()
    {
        if(this.state.siteJobTitle != this.state.aadJobTitle) // fix the job title in the user info list
        {
          SPOQASpinner.Show(strings.UPT_FixSiteTitle);
          RestAPIHelper.FixJobTitleInUserInfoList(this.state.userId, this.props.spHttpClient, this.state.affectedSite, this.state.aadJobTitle, this.SuccessCallBack, this.FailedCallback);            
        }
        else if(this.state.uapJobtitle != this.state.aadJobTitle)
        {
          SPOQASpinner.Show(strings.UPT_FixUserProfileTitle);
          try
          {
            var fixResult =  await RestAPIHelper.FixJobTitleInUserProfile(this.state.affectedUser, this.props.spHttpClient, this.state.affectedSite, this.state.aadJobTitle);
            console.log(fixResult);
          }
          catch(err)
          {
            SPOQAHelper.ShowMessageBar("Error", strings.UPT_FailedUserProfileTitle); 
          }
        }
    }

    public SuccessCallBack()
    {    
        SPOQASpinner.Hide();
        SPOQAHelper.ShowMessageBar("Success", strings.UPT_SuccessUserInfoListTitle); 
    }

    public FailedCallback()
    {
        SPOQASpinner.Hide();
        SPOQAHelper.ShowMessageBar("Error", strings.UPT_FailedUserInfoListTitle);     
    }

    private ResetStatus():void
    {      
      this.state.userId = -1;
      this.state.aadJobTitle = "";
      this.state.uapJobtitle ="";
      this.state.siteJobTitle ="";
      SPOQAHelper.ResetFormStaus();       
    }    
}