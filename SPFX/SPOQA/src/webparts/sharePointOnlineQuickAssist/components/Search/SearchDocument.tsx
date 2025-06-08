import * as React from 'react';
import {  
    DefaultButton,
    TextField,
    Label,
    ComboBox,
    IComboBox,
    IComboBoxOption,
    Panel,
    PanelType,
    Link
  } from 'office-ui-fabric-react/lib/index';
import RestAPIHelper from '../../../Helpers/RestAPIHelper';
import { ISharePointOnlineQuickAssistProps } from '../ISharePointOnlineQuickAssistProps';
import SPOQAHelper from '../../../Helpers/SPOQAHelper';
import SPOQASpinner from '../../../Helpers/SPOQASpinner';
import styles from '../SharePointOnlineQuickAssist.module.scss';
import  { ItemType, ModerationStatusHelper} from '../../../Helpers/ModerationStatusHelper';
import {RemedyHelper} from '../../../Helpers/RemedyHelper';
import * as strings from 'SharePointOnlineQuickAssistWebPartStrings';
import FormsHelper from '../../../Helpers/FormsHelper';
import SearchHelper from '../../../Helpers/SearchHelper';
import CrawlLogGrid from "./CrawlLogGrid";
import ManagedPropertyGrid from "./ManagedPropertyGrid";
import CrawledPropertyGrid from "./CrawledPropertyGrid";
import {Text } from '@microsoft/sp-core-library';
export default class SearchDocumentQA extends React.Component<ISharePointOnlineQuickAssistProps>
{
    public state = {
        affectedDocument:"",
        affectedLibrary:{Title:"",Id:"",RootFolder:"",BaseType:-1},
        siteLibraries:[],         
        affectedSite:this.props.webAbsoluteUrl,
        isLibrary:false,
        siteIsVaild:false,
        isListNoIndex:false,
        isWebNoIndex:false,
        isDraftVersion:false,
        isMissingDisplayForm:false,
        isChecked:false,
        needRemedy:false,       
        remedyStepsShowed:false,
        crawlLogs:[],
        mpPanelOpen:false,
        clPanelOpen:false,
        managedProperties:[],
        cpPanelOpen:false,
        crawledProperties:[]
      };
    private listTitle:string="";
    private resRef= React.createRef<HTMLDivElement>(); 
    private remedyRef = React.createRef<HTMLDivElement>();
    private websNeedFixNoCrawl:string[] =[];
    private redStyle = "color:red";
    private greenStyle = "color:green";    
    private remedySteps:any[] =[]; 
    public render():React.ReactElement<ISharePointOnlineQuickAssistProps>
    {
        return (            
            <div id="SearchDocumentContainer">
                <div className={ styles.row }>
                    <div className={ styles.column }>
                        <div id="QuestionsSection">
                            <TextField
                                    label={strings.AffectedSiteLoadList}
                                    multiline={false}
                                    onChange={(e)=>{let text:any = e.target; this.setState({affectedSite:text.value,siteIsVaild:false, affectedDocument:""}); this.CleanPreviousResult();}}
                                    value={this.state.affectedSite}
                                    required={true}
                                    onKeyDown={(e)=>{if(e.keyCode ===13){
                                        this.LoadLists();
                                        this.CleanPreviousResult();
                                    }}}                          
                            /> 
                            {this.state.siteIsVaild? 
                                <div>
                                    <ComboBox
                                    defaultSelectedKey="-1"
                                    label={strings.SelectList}
                                    allowFreeform
                                    autoComplete="on"
                                    options={this.state.siteLibraries} 
                                    required={true}                    
                                    onChange ={(ev: React.FormEvent<IComboBox>, option?: IComboBoxOption): void => {
                                        this.setState(
                                            {affectedLibrary:
                                                    {Title:option.key, 
                                                    Id:option.data.listId, 
                                                    RootFolder:option.data.rootFolder,                                                  
                                                    BaseType:option.data.baseType}, 
                                                isChecked:false,
                                                isLibrary:option.data.baseType==1,
                                                affectedDocument:""});
                                         this.CleanPreviousResult();  
                                        }} 
                                    />   
                                {this.state.affectedLibrary.Title!=""? 
                                <div><TextField
                                label={strings.AffectedDocument}
                                multiline={false}
                                onChange={(e)=>{
                                    let text:any = e.target;
                                    this.setState({affectedDocument:text.value}); 
                                    this.CleanPreviousResult();
                                    }}
                                value={this.state.affectedDocument}
                                required={true}                                                
                                />
                                        {!this.state.isLibrary?<Label>e.g. {this.props.rootUrl}{this.state.affectedLibrary.RootFolder}/DispForm.aspx?ID=xxx</Label>
                                            :<Label>e.g. {this.props.rootUrl}{this.state.affectedLibrary.RootFolder}/xxxx.xxx</Label>}
                                </div>:null}                
                                </div>: null}
                            </div>
                            <div id="SearchDocumentCheckResultSection">
                                {this.state.isChecked && this.state.siteIsVaild && this.state.affectedLibrary.Title!=""?
                                    <Label>Diagnose result:</Label>                                   
                                    :null}
                                <div style={{marginLeft:20}} id="SearchDocumentCheckResultDiv" ref={this.resRef}></div>
                                {this.state.crawlLogs.length >0?  <Link onClick={e=>{this.setState({clPanelOpen:true});}}  style={{ display: 'block'}}>{strings.SS_Label_CrawlLogs}</Link>:null}                               
                                {this.state.managedProperties.length >0? <Link onClick={e=>{this.setState({mpPanelOpen:true});}}  style={{ display: 'block'}}>{strings.SD_ShowManagedProperties}</Link>:null}                               
                                {this.state.crawledProperties.length >0? <Link onClick={e=>{this.setState({cpPanelOpen:true});}} style={{ display: 'block'}}>{strings.SD_ShowCrawlProperties}</Link>:null}
                            </div>                           
                        <div id="CommandButtonsSection">
                            <DefaultButton
                                text={strings.CheckIssues}
                                style={{ display: 'inline', marginTop: '10px' }}
                                onClick={() => {this.state.siteIsVaild? this.CheckSearchDocument():this.LoadLists();}}
                                />
                            {this.state.needRemedy && !this.state.remedyStepsShowed && this.state.siteIsVaild?
                                <DefaultButton
                                    text={strings.ShowRemedySteps}
                                    style={{ display: 'inline', marginTop: '10px', marginLeft:"10px"}}
                                    onClick={() => {this.ShowRemedySteps();}}
                                />:null}
                        </div>
                        <div id="RemedyStepsDiv" ref={this.remedyRef}>
                                
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
                <Panel
                    headerText={strings.SS_Label_CrawlLogs}
                    isOpen={this.state.clPanelOpen}
                    onDismiss={e=>{this.setState({clPanelOpen:false});}}                   
                    closeButtonAriaLabel="Close"
                    customWidth="100%"
                    type={PanelType.custom}                >
                    <CrawlLogGrid items={this.state.crawlLogs}/>
                </Panel> 
            </div>
        );
    }
    
    public async LoadLists()
    {      
        
        if(this.state.affectedSite == "" || !this.state.affectedSite || !SPOQAHelper.ValidateUrl(this.state.affectedSite))
        {
          SPOQAHelper.ShowMessageBar("Error", strings.UI_NonAffectedSite);          
          return;
        }
        try
        {
            var lists:any = await RestAPIHelper.GetLists(this.props.spHttpClient, this.state.affectedSite);
            let listOptions:IComboBoxOption[] = [];

            // list.BaseType ==1 means this list is a library, otherwise this list is a list
            lists.forEach(list => {              
                    listOptions.push({ 
                        key:list.Title, 
                        text: list.Title,
                        data:{listId:list.Id, baseType:list.BaseType,rootFolder:list.RootFolder.ServerRelativeUrl}});               
            });
            this.setState({siteIsVaild:true, siteLibraries:listOptions});            
        }
        catch(err)
        {
            SPOQAHelper.ShowMessageBar("Error", `${strings.FailedLoadSiteList} ${err}`);
        }        
    }
    
    public async CheckSearchDocument()
    {
       
        if(this.state.affectedLibrary.Title == "" ||this.state.affectedLibrary.Title =="-1")
        {
            SPOQAHelper.ShowMessageBar("Error", strings.PleaseSelectList);
            return;
        }

        if(!this.state.affectedDocument || this.state.affectedDocument =="" || this.state.affectedDocument.trim()=="")
        {
            SPOQAHelper.ShowMessageBar("Error", strings.SD_DocumentPathCanNotBeNull);
            return;
        }

        SPOQAHelper.ResetFormStaus();        
        this.CleanPreviousResult(); 
        let searched:boolean = false;
        SPOQASpinner.Show(`${strings.Checking} ......`);

        // Get crawl logs
        var crawlLogs = await SearchHelper.GetCrawlLogByRest(this.props.spHttpClient,this.state.affectedSite,this.state.affectedDocument);
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
            var needCrawlLogReadPermssionMsg = `<span style="${this.redStyle}">${Text.format(strings.SD_CrawlLackReadLogPermssion, crawlLogPermssionUrl)}.</span><br/>`;
            this.resRef.current.innerHTML += needCrawlLogReadPermssionMsg;            
        }
       
        this.listTitle = this.state.affectedLibrary.Title;
        try
        {
           var searchDocRes = await RestAPIHelper.SearchDocumentByFullPath(this.props.spHttpClient, this.state.affectedSite, this.state.affectedDocument);
           console.log(searchDocRes);
           if(searchDocRes.RowCount >0)
           {    
                // If the document can be searched, then get the ManagedProperties and CrawledProperties for the first row
                var docId = "";
                var docIds = searchDocRes.Table.Rows[0].Cells.filter(e=>e.Key=="DocId");
                if(docIds.length >0)
                {
                    docId = docIds[0].Value;
                }
                
                if(docId)
                {
                   var resCP:any = await SearchHelper.GetCrawledProperties(this.props.spHttpClient, this.state.affectedSite, docId);
                   var resMP:any = await SearchHelper.GetManagedProperties(this.props.spHttpClient, this.state.affectedSite, docId);
                   searched = true;  
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
                    var docIDMissedMsg = `<span style="${this.redStyle}">${Text.format(strings.SD_DocIdIsNull, searchDocRes.RowCount)}</span><br/>`;
                    this.resRef.current.innerHTML += docIDMissedMsg;
                    this.remedySteps.push({
                        message:strings.SD_DocIdIsNullRemedy,
                        url:""});
                }
           }           
        }
        catch(err)
        {
            SPOQAHelper.ShowMessageBar("Error",`${strings.SD_SearchByFullPathException} ${err}`);             
        }

        // Check web no-index
        if(!searched)
        {
            try
            {
                /*var noCrawl = await RestAPIHelper.IsWebNoCrawl(this.props.spHttpClient, this.state.affectedSite);
                this.setState({isWebNoIndex:noCrawl});*/
                await this.CheckWebNoCrawl();   
                this.setState({isChecked:true});         
            }
            catch(err)
            {
                SPOQAHelper.ShowMessageBar("Error",`${strings.SD_IsWebNoCrawlException} ${err}`);                 
            }

            // check library no-index
            try
            {
                var resIsListNoCrawl = await RestAPIHelper.IsListNoCrawl(this.props.spHttpClient, this.state.affectedSite, this.listTitle);
                this.setState({isListNoIndex:resIsListNoCrawl});
                var listNoCrawlMsg = `<span style="${resIsListNoCrawl? this.redStyle:this.greenStyle}" >${resIsListNoCrawl?strings.SD_TheNocrawlEnabledList:strings.SD_TheNocrawlNotEnabledList} ${this.state.affectedLibrary.Title}.</span><br/>`;
                this.resRef.current.innerHTML += listNoCrawlMsg;
                if(resIsListNoCrawl)
                {
                    this.remedySteps.push({
                        message:strings.SD_DectectedNocrawlList,
                        url:`${this.state.affectedSite}/_layouts/15/advsetng.aspx?List={${this.state.affectedLibrary.Id}}`});
                }
            }
            catch(err)
            {
                SPOQAHelper.ShowMessageBar("Error",`${strings.SD_IsListNoCrawlException} ${err}`);                
            }

            // check the list form missed issue
            if(!this.state.isLibrary)
            {
                try
                {
                    var resIsListMissedForm = await RestAPIHelper.IsListMissDisplayForm(this.props.spHttpClient, this.state.affectedSite, this.listTitle);
                    // isMissingDisplayForm
                    this.setState({isMissingDisplayForm:resIsListMissedForm});
                    var listMissedDisplayFormMsg =  `<span style="${resIsListMissedForm? this.redStyle:this.greenStyle}" >${resIsListMissedForm?strings.SD_TheDisplayFormMissed:strings.SD_TheDisplayFormNotMissed} ${this.state.affectedLibrary.Title}</span><br/>`;
                    this.resRef.current.innerHTML += listMissedDisplayFormMsg;
                    if(resIsListMissedForm)
                    {
                        this.remedySteps.push({message:strings.SD_DectectedDisplayFormIsMissing});
                    }
                }
                catch(err)
                {
                    SPOQAHelper.ShowMessageBar("Error",`${strings.SD_IsListMissDisplayFormException} ${err}`);                     
                }
            }

            // Check the approve status, OData__ModerationStatus, https://docs.microsoft.com/en-us/dotnet/api/microsoft.sharepoint.spmoderationstatustype?view=sharepoint-server
            var resNotApprovedItems = await ModerationStatusHelper.GetNotApprovedItems(this.props.spHttpClient, this.state.affectedSite, this.state.affectedLibrary.RootFolder, this.state.affectedDocument);
            if(resNotApprovedItems.success)
               {
                    resNotApprovedItems.items.forEach((item)=>{
                        var notApproveMsg = `<a href='${item.url}'>${item.name}</a>${strings.PC_ApproveStatusIs} ${item.status}.`;
                        this.remedySteps.push({
                            message:notApproveMsg,
                            url:item.parentUrl});
                        this.resRef.current.innerHTML += `<span style="${this.redStyle}" >${notApproveMsg}</span><br/>`;
                });                  
               }      

            // check the draft version issue
            if(resNotApprovedItems.itemType == ItemType.Document || resNotApprovedItems.itemType == ItemType.ListItem)
            {
                try
                {
                    var resIsDraftVersion = await RestAPIHelper.IsDocumentInDraftVersion(this.props.spHttpClient, this.state.affectedSite, this.state.isLibrary, this.listTitle,this.state.affectedDocument);
                    this.setState({isDraftVersion:resIsDraftVersion});
                    var docIsDraftVersionMsg =  `<span style="${resIsDraftVersion? this.redStyle:this.greenStyle}" > ${resIsDraftVersion?strings.PC_DocumentIsInDraft:strings.PC_DocumentIsNotInDraft}</span><br/>`;
                    this.resRef.current.innerHTML += docIsDraftVersionMsg;
                    if(resIsDraftVersion)
                    {
                        this.remedySteps.push({
                            message:strings.PC_DocumentIsInDraft,
                            url:this.props.rootUrl + this.state.affectedLibrary.RootFolder});    
                    }                
                }
                catch(err)
                {
                    SPOQAHelper.ShowMessageBar("Error",`${strings.SD_IsDocumentInDraftVersionException} ${err}`);                
                } 
           }
           else if(resNotApprovedItems.itemType == ItemType.Folder)
           {
                this.resRef.current.innerHTML += `${this.state.affectedDocument} ${strings.SD_FolderSkipDraftCheck}`;
           } 
           else // ItemType is unknow
           {
                SPOQAHelper.ShowMessageBar("Error",`${resNotApprovedItems.error}`);   
           }                        
        }
        
        if(this.remedySteps.length >0)
        {
            this.setState({needRemedy:true,remedyStepsShowed:false});  
        }

        SPOQASpinner.Hide();
    }
    
    // Auto fix has been deprecated, needn't to localize it
    public async FixIssues()
    {
        SPOQAHelper.ResetFormStaus();
        SPOQASpinner.Show("Fix detected document search issues ......");
        let hasError:boolean = false;
        if(this.state.isListNoIndex)
        {
            try
            {
                SPOQASpinner.Show("Fixing list no crawl .......");
                var fixListNoIndexRes = await RestAPIHelper.FixListNoCrawl(this.props.spHttpClient, this.state.affectedSite, this.listTitle);
            }
            catch(err)
            {
                SPOQAHelper.ShowMessageBar("Error",`Get exception when try to check FixListNoCrawl with error message ${err}`);
                hasError = true;
            }
        }
        
        if(this.websNeedFixNoCrawl.length >0)
        {
            for(var index =0; index <this.websNeedFixNoCrawl.length;index++)
            {
                try{
                    var currentWeb = this.websNeedFixNoCrawl[index];
                    SPOQASpinner.Show(`Fixing web no crawl for ${currentWeb} .......`);
                    var fixWebNoIndexRes = await RestAPIHelper.FixWebNoCrawl(this.props.spHttpClient, currentWeb);
                    if(!fixWebNoIndexRes.ok)
                    {
                        SPOQAHelper.ShowMessageBar("Error",`Unable to fix no crawl for ${currentWeb}, please make sure you have permision and <a href="https://docs.microsoft.com/en-us/sharepoint/allow-or-prevent-custom-script">allow script</a> is enabled for the site collection`);
                        hasError = true;
                    }
                }
                catch(err)
                {
                    SPOQAHelper.ShowMessageBar("Error",`Get exception when try to check FixWebNoCrawl with error message ${err}`);
                    hasError = true;
                }
            }
        }

        if(this.state.isDraftVersion)
        {
            try
            {
                SPOQASpinner.Show(`Fixing document draft version .......`);
                var fixDraftVersionRes = await RestAPIHelper.FixDraftVersion(this.props.spHttpClient, this.state.affectedSite, this.state.isLibrary, this.listTitle, this.state.affectedDocument);
            }
            catch(err)
            {
                SPOQAHelper.ShowMessageBar("Error",`Get exception when try to check FixDraftVersion with error message ${err}`);
                hasError = true;
            }
        }

        if(this.state.isMissingDisplayForm)
        {
            try
            {
                SPOQASpinner.Show(`Fixing missed display form .......`);
                var fixMissingDisplayFormRes = await FormsHelper.FixMissDisForm(this.props.spHttpClient, this.state.affectedSite, this.listTitle);
            }
            catch(err)
            {
                SPOQAHelper.ShowMessageBar("Error",`Get exception when try to check FixDraftVersion with error message ${err}`);
                hasError = true;
            }
        }

        if(!hasError)
        {
            SPOQAHelper.ShowMessageBar("Success", `Fixed all detected issues please try to reindex the affected library/site and wait for 20~30 minutes then verify it`);
            this.setState({isChecked:false});            
        }

        SPOQASpinner.Hide();
    }

    public async CheckWebNoCrawl() // check the current web and all parent web
    {
        // this.resRef.current.innerHTML ="";
        this.websNeedFixNoCrawl = [];
        let properties:string[] = ["NoCrawl"];
        let resList:any[] = [];
        var hasParentWeb = true;
        let currentWebUrl = this.state.affectedSite;
        while(hasParentWeb)
        {
            var webInfo = await RestAPIHelper.GetWebInfo(this.props.spHttpClient, currentWebUrl, properties);
            if(webInfo.NoCrawl)
            {               
                this.websNeedFixNoCrawl.push(currentWebUrl);
                this.remedySteps.push({
                    message:`${strings.SD_DectectedNocrawlSite} ${currentWebUrl}.`,
                    url:`${currentWebUrl}/_layouts/15/srchvis.aspx`});
            }
            
            var noIndexMsg = `<span style="${webInfo.NoCrawl? this.redStyle:this.greenStyle}" >${webInfo.NoCrawl?strings.SD_NocrawlEnabledSite:strings.SD_NocrawlNotEnabledSite} ${currentWebUrl}.</span><br/>`;
            this.resRef.current.innerHTML += noIndexMsg;
            currentWebUrl = await RestAPIHelper.GetParentWebUrl(this.props.spHttpClient, currentWebUrl);
            hasParentWeb = currentWebUrl && currentWebUrl!="";
        }       
    }

    private ShowRemedySteps()
    {    
        this.remedyRef.current.innerHTML = RemedyHelper.GetRemedyHtml(this.remedySteps);
        this.setState({remedyStepsShowed:true});   
    }

    private CleanPreviousResult() {
        this.websNeedFixNoCrawl=[];    
        this.resRef.current.innerHTML ="";
        this.remedyRef.current.innerHTML =""; 
        this.remedySteps =[];
        this.setState({isChecked:false, needRemedy:false, crawlLogs:[], managedProperties:[],crawledProperties:[]});
    }

}