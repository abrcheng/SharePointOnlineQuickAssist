import * as React from 'react';
import {  
    PrimaryButton,
    Text,
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
export default class SearchSiteQA extends React.Component<ISharePointOnlineQuickAssistProps>
{
    public state = {
        affectedSite:this.props.webAbsoluteUrl,
        //affectedUser:this.props.currentUser.email,
        isWebThere:false,
        isWebNoIndex:false,
        userPerm:false,
        isinMembers:false,
        GroupId:"",
        isChecked:false
    };

    public render():React.ReactElement<ISharePointOnlineQuickAssistProps>
    {        
        return (
            <div>
                 <div className={ styles.row }>
                    <div className={ styles.column }>
                      <TextField
                            label={strings.SS_Label_AffectedSite}
                            multiline={false}
                            onChange={(e)=>{let text:any = e.target; this.setState({affectedSite:text.value}); this.setState({isChecked:false});}}
                            value={this.state.affectedSite}
                            required={true}                                                
                      />
                        {this.state.affectedSite!="" && this.state.isChecked? 
                            <div id="SearchSiteResultSection">
                                <Label>{strings.SS_DiagnoseResultLabel}</Label>
                                {this.state.isWebThere?<Label style={{"color":"Green",marginLeft:20}}>{strings.SS_FoundSite} {this.state.affectedSite}</Label>:
                                    <Label style={{"color":"Red",marginLeft:20}} >{strings.SS_SiteNoExist1} {this.state.affectedSite} {strings.SS_SiteNoExist2}</Label>}
                                {this.state.isWebThere?
                                <div>
                                {this.state.isWebNoIndex?<Label style={{"color":"Red",marginLeft:20}}>{strings.SS_NoCrawlEnabled} {this.state.affectedSite}</Label>:
                                    <Label style={{"color":"Green",marginLeft:20}}>{strings.SS_SiteSearchable1} {this.state.affectedSite} {strings.SS_SiteSearchable2}</Label>}
                                {this.state.userPerm?<Label style={{"color":"Green",marginLeft:20}}>{strings.SS_HaveAccess}</Label>:
                                    <Label style={{"color":"Red",marginLeft:20}}>{strings.SS_NoAccess}</Label>}
                                </div>:null}
                                {this.state.GroupId?
                                <div>
                                {this.state.isinMembers?<Label style={{"color":"Green",marginLeft:20}}>{strings.SS_InMembers}</Label>:
                                    <Label style={{"color":"Red",marginLeft:20}}>{strings.SS_NotInMembers}</Label>}
                                </div>:null}
                            </div>:null
                        }
                      <div id="CommandButtonsSection">
                        <PrimaryButton
                          text={strings.SS_Label_CheckIssues}
                          style={{ display: 'inline', marginTop: '10px' }}
                          onClick={() => {this.ResetSatus(); this.CheckSiteSearchSettings();}} //When click: Reset banner status & check if the site is searchable
                        />
                        {this.state.isChecked && this.state.isWebThere && (this.state.isWebNoIndex || (this.state.GroupId && !this.state.isinMembers))?
                            <PrimaryButton
                                text={strings.SS_Label_FixIssues}
                                style={{ display: 'inline', marginTop: '10px', marginLeft:"10px"}}
                                onClick={() => {this.FixIssues();}}
                            />:null}
                    </div>
                  </div>
                </div>
            </div>
        );
    }
    
    private async ResetSatus()
    {
      this.state.isWebThere=false;
      this.state.isWebNoIndex=false;
      this.state.userPerm=false;
      this.state.isinMembers=false;
      this.state.GroupId="";
      this.state.isChecked=false;
      SPOQAHelper.ResetFormStaus();
    }

    private async GetJsonResults(JsonStr:string)
    {
      interface ResultsSum {
        queryModifcation: string;
        total: number;
        totalNoDup: number;
      }
      
      var sumObj:ResultsSum = {
        queryModifcation: "string",
        total: JsonStr['PrimaryQueryResult'].RelevantResults.TotalRows,
        totalNoDup: JsonStr['PrimaryQueryResult'].RelevantResults.TotalRowsIncludingDuplicates
      };
      console.log(sumObj);

      return sumObj;

    }

    private async getUserIDByEmail(email:string,siteUrl:string):Promise<number>
    {
        var url = `${siteUrl}/_api/web/siteusers?$filter=Email eq '${email}'`;
        var userData:any = await RestAPIHelper.GetQueryUser(siteUrl,this.props.spHttpClient);
        return userData.value[0].Id;
    }

    

    public async CheckSiteSearchSettings()
    {
        this.setState({isChecked:false});
        SPOQASpinner.Show(`${strings.SS_Message_Checking}`);
        try
        {
          let url:URL = new URL(this.state.affectedSite);
          let rootSiteUrl = `${url.protocol}//${url.hostname}`;
          console.log(rootSiteUrl);
          var siteSearch = await RestAPIHelper.GetSerchResults(this.props.spHttpClient, rootSiteUrl, this.state.affectedSite, "Site");
          console.log(siteSearch);

          var sum = await this.GetJsonResults(siteSearch);

          if(sum.total == 0)
          {
            if(sum.totalNoDup == 0)
            {
              console.log(`${strings.SS_Message_NoSearchResult}`);
              SPOQAHelper.ShowMessageBar("Error", `${strings.SS_Message_NoSearchResult}`); 

              //Site is not searchable. Proceed to check more
              //1. Check if the site exists
              //2. Check if the site is crawled/indexed
              //3. Check if the user has permissions
              //4. Try searching the site with a different keyword

              //Check if the site exists
              try{
                var webInfo = await RestAPIHelper.GetWeb(this.props.spHttpClient, this.state.affectedSite);
                this.setState({isWebThere:webInfo});
              }
              catch(err)
              {
                SPOQAHelper.ShowMessageBar("Error",`${strings.SS_Ex_GetWebError} ${err}`);
                return;
              }
              if(webInfo)
              {
                //Check if the site is crawled/indexed
                try
                {
                  var noCrawl = await RestAPIHelper.IsWebNoCrawl(this.props.spHttpClient, this.state.affectedSite);
                  this.setState({isWebNoIndex:noCrawl});
                }
                catch(err)
                {
                  SPOQAHelper.ShowMessageBar("Error",`${strings.SS_Ex_IsWebNoCrawlError} ${err}`);
                  return;
                }

                //Check if the user has permissions
                try
                {
                  //Get User Login Id by Email
                  //var userLoginId = this.getUserIDByEmail(this.state.affectedUser, this.state.affectedSite);

                  var userInfoSite = await RestAPIHelper.GetUserFromUserInfoList(this.props.currentUser.email, this.props.spHttpClient, this.state.affectedSite);
                  //↑ Get error if user email contains "'"
                  if(userInfoSite != null)
                  {
                    var permRes = await RestAPIHelper.GetUserPermissions(this.props.currentUser.email, this.props.spHttpClient, this.state.affectedSite);
                    console.log(permRes);
                    this.state.userPerm = permRes;

                    //Check if the site is a group site
                    var groupid = await RestAPIHelper.GetSiteGroupId(this.props.spHttpClient, this.props.ctx, this.state.affectedSite);
                    console.log(groupid);
                    
                    if(groupid == "00000000-0000-0000-0000-000000000000")
                    {
                      this.state.GroupId = "";
                    }
                    else
                    {
                      this.state.GroupId = groupid;
                    }
                    
                    if(this.state.GroupId)
                    {
                      //Get the group members
                      var memberInfo = await GraphAPIHelper.GetGroupMembers(groupid, this.props.msGraphClient);
                      console.log(memberInfo);
                      if(memberInfo.length > 0)
                      {
                        for(var i=0;i<memberInfo.length;i++)
                        {
                          console.log(memberInfo[i]);
                          if(memberInfo[i]['mail'] == this.props.currentUser.email) //Check if current user is in members
                          {
                            this.state.isinMembers = true;
                            break;
                          }
                        }
                      }
                      else
                      {
                        this.state.isinMembers = false;
                      }                      
                    }

                  }
                  else
                  {
                    this.state.userPerm = false;
                  }
                }
                catch(err)
                {
                  SPOQAHelper.ShowMessageBar("Error",`${strings.SS_Ex_GetUserInfoError} ${err}`);
                  return;
                }
              }
            }
            else
            {
              console.log(`${strings.SS_Message_ResultDuplicate}`);
              SPOQAHelper.ShowMessageBar("Warning", `${strings.SS_Message_ResultDuplicate}`); 
            }
          }
          else
          {
            console.log(`${strings.SS_Message_SiteSearchable}`);
            SPOQAHelper.ShowMessageBar("Success", `${strings.SS_Message_SiteSearchable}`); 
            SPOQASpinner.Hide();
            this.setState({isChecked:false});
            return;
          }
        }
        catch(err)
        {
           this.forceUpdate();
           console.log(`Error`);
        }

        SPOQASpinner.Hide();
        this.setState({isChecked:true});
    }

    public async FixIssues()
    {
        SPOQAHelper.ResetFormStaus();
        SPOQASpinner.Show(`${strings.SS_Message_FixSite}`);
        let hasError:boolean = false;
        
        if(this.state.isWebNoIndex)
        {
            try{
                await RestAPIHelper.FixWebNoCrawl(this.props.spHttpClient, this.state.affectedSite);
            }
            catch(err)
            {
                SPOQAHelper.ShowMessageBar("Error",`${strings.SS_Ex_FixWebNoCrawlError} ${err}`);
                hasError = true;
            }
        }

        if(this.state.GroupId && !this.state.isinMembers)
        {
            try
            {
                var addUserinMembers = await GraphAPIHelper.AddUserinMembers(this.state.GroupId, this.props.msGraphClient, this.props.currentUser.email);
            }
            catch(err)
            {
                SPOQAHelper.ShowMessageBar("Error",`${strings.SS_Ex_AddUserInMembersError} ${err}`);
                hasError = true;
            }
        }

        if(!hasError)
        {
            SPOQAHelper.ShowMessageBar("Success", `${strings.SS_Message_FxiedAll}`);
            this.setState({isChecked:false});
        }

        SPOQASpinner.Hide();

    }
}
