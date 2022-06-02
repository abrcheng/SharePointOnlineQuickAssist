import * as React from 'react';
import {  
    PrimaryButton,    
    TextField,
    Label,
    Panel,
    PanelType,
    Link
  } from 'office-ui-fabric-react/lib/index';
import GraphAPIHelper from '../../../Helpers/GraphAPIHelper';  
import RestAPIHelper from '../../../Helpers/RestAPIHelper';
import SPOQASpinner from '../../../Helpers/SPOQASpinner';
import SPOQAHelper from '../../../Helpers/SPOQAHelper';
import { ISharePointOnlineQuickAssistProps } from '../ISharePointOnlineQuickAssistProps';
import styles from '../SharePointOnlineQuickAssist.module.scss';
import * as strings from 'SharePointOnlineQuickAssistWebPartStrings';
import {RemedyHelper} from '../../../Helpers/RemedyHelper';
import FormsHelper from '../../../Helpers/FormsHelper';
import SearchHelper from '../../../Helpers/SearchHelper';
import {Text} from '@microsoft/sp-core-library';
import CrawlLogGrid from "./CrawlLogGrid";
import ManagedPropertyGrid from "./ManagedPropertyGrid";
import CrawledPropertyGrid from "./CrawledPropertyGrid";

export default class SearchSiteQA extends React.Component<ISharePointOnlineQuickAssistProps>
{
    public state = {
        affectedSite:this.props.webAbsoluteUrl,
        //affectedUser:this.props.currentUser.email,
        isWebThere:false,
        isWebNoIndex:true,
        userPerm:false,
        isinMembers:false,
        GroupId:"",
        isChecked:false,
        crawlLogs:[],
        crawlLogerror:"",
        mpPanelOpen:false,
        managedProperties:[],
        cpPanelOpen:false,
        crawledProperties:[]
    };

    private remedySteps =[]; 
    private remedyRef = React.createRef<HTMLDivElement>();

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
                            onKeyDown={(e)=>{if(e.keyCode ===13){this.ResetSatus(); this.CheckSiteSearchSettings();}}}
                      />
                        {this.state.affectedSite!="" && this.state.isChecked? 
                            <div id="SearchSiteResultSection">
                                <Label>{strings.SS_DiagnoseResultLabel}</Label>
                                {this.state.isWebThere?<Label style={{"color":"Green",marginLeft:20}}>{Text.format(strings.SS_FoundSite, this.state.affectedSite)}</Label>:
                                    <Label style={{"color":"Red",marginLeft:20}} >{Text.format(strings.SS_SiteNoExist1,this.state.affectedSite)}</Label>}
                                {this.state.isWebThere?
                                <div>
                                {this.state.userPerm?<div><Label style={{"color":"Green",marginLeft:20}}>{strings.SS_HaveAccess}</Label>
                                    {this.state.isWebNoIndex?<Label style={{"color":"Red",marginLeft:20}}>{Text.format(strings.SS_NoCrawlEnabled1,this.state.affectedSite)}</Label>:
                                    <Label style={{"color":"Green",marginLeft:20}}>{Text.format(strings.SS_SiteIndexEnabled1, this.state.affectedSite)}</Label>}
                                    {this.state.GroupId?
                                    <div>
                                    {this.state.isinMembers?<Label style={{"color":"Green",marginLeft:20}}>{strings.SS_InMembers}</Label>:
                                        <Label style={{"color":"Red",marginLeft:20}}>{strings.SS_NotInMembers}</Label>}
                                    </div>:null}
                                  </div>:<Label style={{"color":"Red",marginLeft:20}}>{strings.SS_NoAccess}</Label>}
                                </div>:null}
                            </div>:null
                        }
                        {this.state.affectedSite!="" && this.state.isChecked? 
                            <div>
                              <div id="FixSuggestionsSection" ref={this.remedyRef}>
                              </div>
                              <Label>{strings.SS_Message_WaitAfterFix}</Label>
                            </div>:null
                        }
                        {this.state.crawlLogs.length >0? <CrawlLogGrid items={this.state.crawlLogs}/>:(this.state.crawlLogerror.length >0?<Label style={{"color":"Red"}}>{this.state.crawlLogerror}</Label>:null)}
                        {this.state.managedProperties.length >0? <Link onClick={e=>{this.setState({mpPanelOpen:true});}}  style={{ display: 'block'}}>{strings.SD_ShowManagedProperties}</Link>:null}                               
                        {this.state.crawledProperties.length >0? <Link onClick={e=>{this.setState({cpPanelOpen:true});}} style={{ display: 'block'}}>{strings.SD_ShowCrawlProperties}</Link>:null}
                      <div id="CommandButtonsSection">
                        <PrimaryButton
                          text={strings.SS_Label_CheckIssues}
                          style={{ display: 'inline', marginTop: '10px' }}
                          onClick={() => {this.ResetSatus(); this.CheckSiteSearchSettings();}} //When click: Reset banner status & check if the site is searchable
                        />
                        {this.state.isChecked && this.state.isWebThere && (this.state.isWebNoIndex || (this.state.GroupId && !this.state.isinMembers))?
                            //<PrimaryButton
                            //    text={strings.SS_Label_FixIssues}
                            //    style={{ display: 'inline', marginTop: '10px', marginLeft:"10px"}}
                            //    onClick={() => {this.FixIssues();}}
                            ///>
                            null:null}
                    </div>
                  </div>
                </div>
                <Panel
                    headerText={strings.SD_ManagedProperties}
                    isOpen={this.state.mpPanelOpen}
                    onDismiss={e=>{this.setState({mpPanelOpen:false});}}                   
                    closeButtonAriaLabel="Close"
                    customWidth='500px'
                    type={PanelType.custom}
                >
                    <ManagedPropertyGrid items={this.state.managedProperties}></ManagedPropertyGrid>
                </Panel> 
                <Panel
                    headerText={strings.SD_CrawlProperties}
                    isOpen={this.state.cpPanelOpen}
                    onDismiss={e=>{this.setState({cpPanelOpen:false});}}                   
                    closeButtonAriaLabel="Close"
                    customWidth='500px'
                    type={PanelType.custom}
                >
                    <CrawledPropertyGrid items={this.state.crawledProperties}></CrawledPropertyGrid>
                </Panel> 
            </div>
        );
    }
    
    private async ResetSatus()
    {
      this.setState({isWebThere:false, isWebNoIndex:false, userPerm:false, isinMembers:false, GroupId:"", isChecked:false, crawlLogs:[], crawlLogerror:"", managedProperties:[],crawledProperties:[]});
      this.remedyRef.current.innerHTML =""; // Clean the RemedyStepsDiv
      SPOQAHelper.ResetFormStaus();
      SPOQASpinner.Hide();
    }

    private ShowRemedySteps()
    {    
        this.remedyRef.current.innerHTML = RemedyHelper.GetRemedyHtml(this.remedySteps);
        this.setState({remedyStepsShowed:true});   
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
        this.remedySteps =[]; 
        SPOQASpinner.Show(`${strings.SS_Message_Checking}`);        
        
        // Get crawl logs
        var crawlLogs = await SearchHelper.GetCrawlLogByRest(this.props.spHttpClient,this.state.affectedSite,this.state.affectedSite);
        if(crawlLogs._ObjectType_ == "SP.SimpleDataTable")
        {
            crawlLogs.Rows.forEach(e=>{
                e.TimeStamp = new Date(parseInt(e.TimeStampUtc.substring(6))).toISOString();
                e.IsDeleted = e.IsDeleted.toString();
            });
            this.setState({crawlLogs:crawlLogs.Rows}); 
        }
        else if(crawlLogs.length>0 && crawlLogs[0].ErrorInfo)
        {
            // need crawl log permssion https://abrcheng-admin.sharepoint.cn/_layouts/15/searchadmin/crawllogreadpermission.aspx
            var crawlLogPermssionUrl = this.props.rootUrl.replace(".sharepoint","-admin.sharepoint") + "/_layouts/15/searchadmin/crawllogreadpermission.aspx";
            this.state.crawlLogerror = Text.format(strings.SS_CrawlLackReadLogPermssion, crawlLogPermssionUrl);
        }
       
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
                //Check if the user has permissions
                try
                {
                  //Get User Login Id by Email
                  //var userLoginId = this.getUserIDByEmail(this.state.affectedUser, this.state.affectedSite);

                  //var userInfoSite = await RestAPIHelper.GetUserFromUserInfoList(this.props.currentUser.email, this.props.spHttpClient, this.state.affectedSite);
                  //↑ Get error if user email contains "'"
                  //if(userInfoSite != null)
                  //{
                    var permRes = await RestAPIHelper.GetUserReadPermissions(this.props.currentUser.email, this.props.spHttpClient, this.state.affectedSite);
                    console.log(permRes);
                    this.state.userPerm = permRes;
                    if(!permRes)
                    {
                      this.remedySteps.push({
                        message:`${strings.SS_Message_CheckPermissions} ${this.state.affectedSite}`});
                    }
                    else
                    {
                      //Check if the site is crawled/indexed
                      try
                      {
                        var hasParentWeb = true;
                        let currentWebUrl = this.state.affectedSite;
                        while(hasParentWeb)
                        {
                            var noCrawl = await RestAPIHelper.IsWebNoCrawl(this.props.spHttpClient, currentWebUrl);
                            if(noCrawl)
                            {
                              this.setState({isWebNoIndex:noCrawl});  
                              this.remedySteps.push({
                                message:`${strings.SS_Message_SearchAndOffline} ${currentWebUrl}`,
                                url:`${currentWebUrl}/_layouts/15/srchvis.aspx`
                              });
                            }
                            
                            currentWebUrl = await RestAPIHelper.GetParentWebUrl(this.props.spHttpClient, currentWebUrl);
                            hasParentWeb = currentWebUrl && currentWebUrl!="";
                        }     
                      }
                      catch(err)
                      {
                        SPOQAHelper.ShowMessageBar("Error",`${strings.SS_Ex_IsWebNoCrawlError} ${err}`);
                        return;
                      }

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
                          this.remedySteps.push({
                            message:`${strings.SS_Message_AddInMembers}`});
                        }                      
                      }

                    }
                  //}
                  //else
                  //{
                  //  this.state.userPerm = false;
                  //  this.remedySteps.push({
                  //    message:`${strings.SS_Message_CheckPermissions} ${this.state.affectedSite}`});
                  //}
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
            console.log(siteSearch.PrimaryQueryResult.RelevantResults);
            this.setState({isChecked:false});
            
            // If the site can be searched, then get the ManagedProperties and CrawledProperties for the first row
            var docId = "";
            var docIds = siteSearch.PrimaryQueryResult.RelevantResults.Table.Rows[0].Cells.filter(e=>e.Key=="DocId");
            if(docIds.length >0)
            {
                docId = docIds[0].Value;
            }
            if(docId)
            {
               var resCP:any = await SearchHelper.GetCrawledProperties(this.props.spHttpClient, this.state.affectedSite, docId);
               var resMP:any = await SearchHelper.GetManagedProperties(this.props.spHttpClient, this.state.affectedSite, docId);
               SPOQAHelper.ShowMessageBar("Success", strings.SD_DocumentCanBeSearched);  
               var mps = [];
               if(resMP.PrimaryQueryResult.RelevantResults.RowCount >0)
               {
                    resMP.PrimaryQueryResult.RelevantResults.Table.Rows[0].Cells.forEach(e=>{mps.push({Name:e.Key,Value:e.Value});});
               }                  
              
               // res.PrimaryQueryResult.RefinementResults.Refiners[0].Entries[0].RefinementName
              var cps = [];
              if(resCP.PrimaryQueryResult.RefinementResults.Refiners.length >0 && resCP.PrimaryQueryResult.RefinementResults.Refiners[0].Entries.length >0)
              {
                resCP.PrimaryQueryResult.RefinementResults.Refiners[0].Entries.forEach(e=>{cps.push({Name:e.RefinementName});});                    
              }
              
              this.setState({managedProperties:mps, crawledProperties:cps});

              SearchHelper.CallOtherDiagnosticsAPIS(this.props.spHttpClient, this.state.affectedSite, docId);
            }
            else
            {
                // DocID is null will causd unexpected issue 
                var docIDMissedMsg = `<span style="color:red">${Text.format(strings.SD_DocIdIsNull, siteSearch.RowCount)}</span><br/>`;
                //this.resRef.current.innerHTML += docIDMissedMsg;
                this.remedySteps.push({
                    message:docIDMissedMsg,
                    url:""});
            }
            
            SPOQASpinner.Hide();
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
        this.ShowRemedySteps();
    }
    /*
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

    }*/
}
