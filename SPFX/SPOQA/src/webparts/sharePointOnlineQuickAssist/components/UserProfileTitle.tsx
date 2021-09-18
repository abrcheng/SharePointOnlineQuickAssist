import * as React from 'react';
import {  
    PrimaryButton,
    Text,
    TextField,
    Label
  } from 'office-ui-fabric-react/lib/index';
import GraphAPIHelper from '../../Helpers/GraphAPIHelper';  
import RestAPIHelper from '../../Helpers/RestAPIHelper';
import SPOQASpinner from '../../Helpers/SPOQASpinner';
import SPOQAHelper from '../../Helpers/SPOQAHelper';
import { ISharePointOnlineQuickAssistProps } from './ISharePointOnlineQuickAssistProps';
export default class UserProfileTitleQA extends React.Component<ISharePointOnlineQuickAssistProps>
{
    public state = {
        affectedUser: "",
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
                <TextField
                        label="Affected Site:"
                        multiline={false}
                        onChange={(e)=>{let text:any = e.target; this.setState({affectedSite:text.value});}}
                        value={this.state.affectedSite}
                        required={true}                        
                  /> 
                <TextField
                        label="Affected User:"
                        multiline={false}
                        onChange={(e)=>{let text:any = e.target; this.setState({affectedUser:text.value});}}
                        value={this.state.affectedUser}
                        required={true}                                                
                  />                  
                  {this.state.aadJobTitle != ""? <Label>JobTitle from AAD is <span style={{"color":"Green"}}>{this.state.aadJobTitle}</span></Label> : null}
                  {this.state.aadJobTitle != "" && this.state.userId && this.state.userId !=-1? <Label>JobTitle from User Profile is <span style={this.state.uapJobtitle != this.state.aadJobTitle? {"color":"Red"}:{"color":"Green"}}>{this.state.uapJobtitle}</span></Label>: null}
                  {this.state.aadJobTitle != "" && this.state.userId && this.state.userId !=-1?<Label>JobTitle from site user info list is <span style={this.state.siteJobTitle != this.state.aadJobTitle? {"color":"Red"}:{"color":"Green"}}>{this.state.siteJobTitle}</span></Label>: null}

                  <PrimaryButton
                      text="Check Job Title"
                      style={{ display: 'inline', marginTop: '10px' }}
                      onClick={() => {this.CheckUserProfileTitle();}}
                    />
                    
                    { (this.state.siteJobTitle != this.state.aadJobTitle || this.state.aadJobTitle != this.state.uapJobtitle) && this.state.userId && this.state.userId !=-1 ? 
                    <PrimaryButton
                      text="Fix It"
                      style={{ display: 'inline', marginTop: '10px', marginLeft:"20px"}}
                      hidden={this.state.siteJobTitle == this.state.aadJobTitle && this.state.uapJobtitle == this.state.aadJobTitle}
                      onClick={() => {this.FixJobTitle();}}
                    />: null}
            </div>
        );
    }

    public async CheckUserProfileTitle()
    {
        // reset status 
        this.ResetStatus();
        
        if(this.state.affectedSite =="" || !this.state.affectedSite || !SPOQAHelper.ValidateUrl(this.state.affectedSite))
        {
          SPOQAHelper.ShowMessageBar("Error", "Affected site can't be null or invalid!");          
          return;
        }

        if(this.state.affectedUser =="" || !this.state.affectedUser || !SPOQAHelper.ValidateEmail(this.state.affectedUser))
        {
          SPOQAHelper.ShowMessageBar("Error", "Affected user can't be null or invalid!");
          return;
        }      

        SPOQASpinner.Show("Checking ...");
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
              SPOQAHelper.ShowMessageBar("Error", "Failed to get the user/job title from AAD!");              
           }
           else if(this.state.userId ==-1)
           {
              SPOQAHelper.ShowMessageBar("Error", "Failed to get the user from the site!");
           }
        }

        SPOQASpinner.Hide();
    }    

    public async FixJobTitle()
    {
        if(this.state.siteJobTitle != this.state.aadJobTitle) // fix the job title in the user info list
        {
          SPOQASpinner.Show("Fixing JobTitle in the site ...");
          RestAPIHelper.FixJobTitleInUserInfoList(this.state.userId, this.props.spHttpClient, this.state.affectedSite, this.state.aadJobTitle, this.SuccessCallBack, this.FailedCallback);            
        }
        else if(this.state.uapJobtitle != this.state.aadJobTitle)
        {
          SPOQASpinner.Show("Fixing JobTitle user profile ...");
          try
          {
            var fixResult =  await RestAPIHelper.FixJobTitleInUserProfile(this.state.affectedUser, this.props.spHttpClient, this.state.affectedSite, this.state.aadJobTitle);
            console.log(fixResult);
          }
          catch(err)
          {
            SPOQAHelper.ShowMessageBar("Error", "Fix JobTitle in the user profile failed."); 
          }
        }
    }

    public SuccessCallBack()
    {    
        SPOQASpinner.Hide();
        SPOQAHelper.ShowMessageBar("Success", "Fix JobTitle in the user info list done, please recheck for verfiying it."); 
    }

    public FailedCallback()
    {
        SPOQASpinner.Hide();
        SPOQAHelper.ShowMessageBar("Error", "Fix JobTitle in the user info list failed.");     
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