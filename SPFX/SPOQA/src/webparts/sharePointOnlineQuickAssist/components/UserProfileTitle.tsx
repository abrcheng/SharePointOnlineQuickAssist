import * as React from 'react';
import {  
    PrimaryButton,
    Text,
    TextField,
    Label
  } from 'office-ui-fabric-react/lib/index';
import GraphAPIHelper from '../../Helpers/GraphAPIHelper';  
import RestAPIHelper from '../../Helpers/RestAPIHelper';
import { ISharePointOnlineQuickAssistProps } from './ISharePointOnlineQuickAssistProps';
export default class UserProfileTitleQA extends React.Component<ISharePointOnlineQuickAssistProps>
{
    public state = {
        affectedUser: "",
        aadJobTitle:"",
        uapJobtitle:"",
        siteJobTitle:"",
        userId:-1
      };

    public render():React.ReactElement<ISharePointOnlineQuickAssistProps>
    {
       
        return (
            <div>
                <TextField
                        label="Affected User:"
                        multiline={true}
                        onChange={(e)=>{let text:any = e.target; this.setState({affectedUser:text.value});}}
                        value={this.state.affectedUser}
                        required={true}                        
                  />                  
                  {this.state.aadJobTitle != ""? <Label>JobTitle from AAD is <span style={{"color":"Green"}}>{this.state.aadJobTitle}</span></Label> : null}
                  {this.state.aadJobTitle != ""? <Label>JobTitle from User Profile is <span style={this.state.uapJobtitle != this.state.aadJobTitle? {"color":"Red"}:{"color":"Green"}}>{this.state.uapJobtitle}</span></Label>: null}
                  {this.state.aadJobTitle != ""?<Label>JobTitle from site user info list is <span style={this.state.siteJobTitle != this.state.aadJobTitle? {"color":"Red"}:{"color":"Green"}}>{this.state.siteJobTitle}</span></Label>: null}

                  <PrimaryButton
                      text="Check Job Title"
                      style={{ display: 'inline', marginTop: '10px' }}
                      onClick={() => {this.CheckUserProfileTitle();}}
                    />
                    
                    { this.state.siteJobTitle != this.state.aadJobTitle || this.state.aadJobTitle != this.state.uapJobtitle? 
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
        console.log("Start to CheckUserProfileTitle");
        var userInfoAAD = await GraphAPIHelper.GetUserInfo(this.state.affectedUser, this.props.msGraphClient);        
        console.log(`Job title from AAD is ${userInfoAAD.jobTitle}`);
        this.setState({aadJobTitle:userInfoAAD.jobTitle});
        var userInfoUAP = await RestAPIHelper.GetUserInfoFromUserProfile(this.state.affectedUser,this.props.spHttpClient, this.props.webAbsoluteUrl);
        console.log(`Job title from UAP is ${userInfoUAP.Title}`);
        this.setState({uapJobtitle:userInfoUAP.Title});
        var userInfoSite = await RestAPIHelper.GetUserFromUserInfoList(this.state.affectedUser, this.props.spHttpClient, this.props.webAbsoluteUrl);
        console.log(`Job title from user info list is ${userInfoSite.JobTitle}`);
        this.setState({siteJobTitle:userInfoSite.JobTitle});    
        this.state.userId = userInfoSite.Id;    
        console.log("ended CheckUserProfileTitle");
    }    

    public async FixJobTitle()
    {
        if(this.state.siteJobTitle != this.state.aadJobTitle) // fix the job title in the user info list
        {
          var updatedUserItem = await RestAPIHelper.FixJobTitleInUserInfoList(this.state.userId, this.props.spHttpClient, this.props.webAbsoluteUrl, this.state.aadJobTitle, this.SuccessCallBack, this.FailedCallback);            
        }
    }

    public SuccessCallBack()
    {
        alert("FixJobTitleInUserInfoList done.");
        this.CheckUserProfileTitle();
    }

    public FailedCallback()
    {
        alert("FixJobTitleInUserInfoList failed.");
    }
}