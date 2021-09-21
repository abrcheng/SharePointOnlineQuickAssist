import * as React from 'react';
import { PrimaryButton, TextField,Label} from 'office-ui-fabric-react/lib/index';
import GraphAPIHelper from '../../Helpers/GraphAPIHelper';  
import RestAPIHelper from '../../Helpers/RestAPIHelper';
import SPOQASpinner from '../../Helpers/SPOQASpinner';
import SPOQAHelper from '../../Helpers/SPOQAHelper';
import { ISharePointOnlineQuickAssistProps } from './ISharePointOnlineQuickAssistProps';

export default class UserProfilePhotoQA extends React.Component<ISharePointOnlineQuickAssistProps>
{
    public state = {
        affectedUser: "",
        aadUserPhotoUrl:"",
        uapUserPhotoUrl:""        
      };
    public mySiteHost:string ="";
    public render():React.ReactElement<ISharePointOnlineQuickAssistProps>
    {        
        this.mySiteHost = this.props.webAbsoluteUrl.replace(this.props.webUrl == "/"?"":this.props.webUrl,"").replace(".sharepoint.com", "-my.sharepoint.com");
        return (
            <div> 
                  
                  <TextField
                        label="Affected User:"
                        multiline={false}
                        onChange={(e)=>{let text:any = e.target; this.setState({affectedUser:text.value});}}
                        value={this.state.affectedUser}
                        required={true}                                                
                  />                  
                  {this.state.aadUserPhotoUrl!=""? <Label>Picture from AAD:</Label>:null}
                  <img src={this.state.aadUserPhotoUrl} />   
                  {this.state.aadUserPhotoUrl!=""?<Label>Picture from user profile:</Label>:null}
                  {this.state.uapUserPhotoUrl!=""? <img src={`${this.props.webAbsoluteUrl}/_layouts/15/userphoto.aspx?username=${this.state.affectedUser}`} />:null}
                  <PrimaryButton
                      text="Check UserProfile Photo"
                      style={{ display: 'block', marginTop: '10px' }}
                      onClick={() => {this.CheckUserPhoto();}}
                  />
            </div>
        );
    }
    
    private async CheckUserPhoto()
    {
        SPOQAHelper.ResetFormStaus();       
        try{
            var userPhoto = await GraphAPIHelper.GetUserPhoto(this.state.affectedUser, this.props.msGraphClient);            
            const blobUrl = window.URL.createObjectURL(userPhoto);
            this.setState({aadUserPhotoUrl:blobUrl});            
            console.log("GraphAPIHelper.GetUserPhoto done");                        
            
            try
            {
                var userInfoUAP = await RestAPIHelper.GetUserInfoFromUserProfile(this.state.affectedUser,this.props.spHttpClient, this.props.webAbsoluteUrl);
                console.log(`Picture URL from UAP is ${userInfoUAP.PictureUrl}`);
                if(userInfoUAP.PictureUrl && userInfoUAP.PictureUrl!="")
                {
                    this.setState({uapUserPhotoUrl:userInfoUAP.PictureUrl});
                }
                SPOQAHelper.ShowMessageBar("Success", "Photos are loaded, but if they are mismatched, please follow this article https://github.com/abrcheng/SharePointOnlineQuickAssist/tree/main/KBs/SyncPhotoFromADToSPO");
            }
            catch(err)
            {
                SPOQAHelper.ShowMessageBar("Error", "Failed to get the user photo from the user profie! Please consider to follow this article https://github.com/abrcheng/SharePointOnlineQuickAssist/tree/main/KBs/SyncPhotoFromADToSPO");
            }
        }
        catch(err)
        {
            SPOQAHelper.ShowMessageBar("Error", "Failed to get the user photo from the AAD! please consider to set the user photo in AAD https://docs.microsoft.com/en-us/azure/active-directory/fundamentals/active-directory-users-profile-azure-portal");
            this.setState({aadUserPhotoUrl:"", uapUserPhotoUrl:""});              
        }        
    }
}