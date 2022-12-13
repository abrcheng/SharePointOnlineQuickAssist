import * as React from 'react';
import { DefaultButton, TextField,Label} from 'office-ui-fabric-react/lib/index';
import GraphAPIHelper from '../../../Helpers/GraphAPIHelper';  
import RestAPIHelper from '../../../Helpers/RestAPIHelper';
import SPOQASpinner from '../../../Helpers/SPOQASpinner';
import SPOQAHelper from '../../../Helpers/SPOQAHelper';
import * as strings from 'SharePointOnlineQuickAssistWebPartStrings';
import { ISharePointOnlineQuickAssistProps } from '../ISharePointOnlineQuickAssistProps';
import styles from '../SharePointOnlineQuickAssist.module.scss';
export default class UserProfilePhotoQA extends React.Component<ISharePointOnlineQuickAssistProps>
{
    public state = {
        affectedUser: this.props.currentUser.loginName,
        aadUserPhotoUrl:"",
        uapUserPhotoUrl:""        
      };
    public mySiteHost:string ="";
    public render():React.ReactElement<ISharePointOnlineQuickAssistProps>
    {        
        this.mySiteHost = this.props.webAbsoluteUrl.replace(this.props.webUrl == "/"?"":this.props.webUrl,"").replace(".sharepoint.com", "-my.sharepoint.com");
        return (
            <div> 
                <div className={ styles.row }>
                    <div className={ styles.column }>
                        <TextField
                                label={strings.AffectedUser}
                                multiline={false}
                                onChange={(e)=>{let text:any = e.target; this.setState({affectedUser:text.value});}}
                                value={this.state.affectedUser}
                                required={true}                                                
                        />    
                        <Label>e.g. John@contoso.com </Label>                 
                        {this.state.aadUserPhotoUrl!=""? <div><Label>{strings.UPP_PhotoAAD}</Label> <img src={this.state.aadUserPhotoUrl} /> </div>:null}                         
                        {this.state.aadUserPhotoUrl!=""?<Label>{strings.UPP_PhotoUserProfile}</Label>:null}
                        {this.state.uapUserPhotoUrl!=""? <img src={`${this.props.webAbsoluteUrl}/_layouts/15/userphoto.aspx?username=${this.state.affectedUser}`} />:null}
                        <DefaultButton
                            text={strings.UI_CheckIssueforUser}
                            style={{ display: 'block', marginTop: '10px' }}
                            onClick={() => {this.CheckUserPhoto();}}
                        />
                    </div>
              </div>  
            </div>
        );
    }
    
    private async CheckUserPhoto()
    {
        SPOQAHelper.ResetFormStaus();      
        
        if(this.state.affectedUser =="" || !this.state.affectedUser || !SPOQAHelper.ValidateEmail(this.state.affectedUser))
        {
          SPOQAHelper.ShowMessageBar("Error", strings.UI_NonAffectedUser);
          return;
        }      
        try{
            var userPhoto = await GraphAPIHelper.GetUserPhoto(this.state.affectedUser, this.props.msGraphClient);            
            const blobUrl = window.URL.createObjectURL(userPhoto);
            this.setState({aadUserPhotoUrl:blobUrl});            
            console.log("GraphAPIHelper.GetUserPhoto done");                        
            
            try
            {
                var userInfoUAP = await RestAPIHelper.GetUserInfoFromUserProfile(this.state.affectedUser,this.props.spHttpClient, this.props.webAbsoluteUrl);
                console.log(`${strings.UPP_PhotoURL} ${userInfoUAP.PictureUrl}`);
                if(userInfoUAP.PictureUrl && userInfoUAP.PictureUrl!="")
                {
                    this.setState({uapUserPhotoUrl:userInfoUAP.PictureUrl});
                }
                SPOQAHelper.ShowMessageBar("Success", strings.UPP_PhotoSuccess);
            }
            catch(err)
            {
                SPOQAHelper.ShowMessageBar("Error", strings.UPP_PhotoFailed);
            }
        }
        catch(err)
        {
            SPOQAHelper.ShowMessageBar("Error", strings.UPP_NonAADPhoto);
            this.setState({aadUserPhotoUrl:"", uapUserPhotoUrl:""});              
        }        
    }
}