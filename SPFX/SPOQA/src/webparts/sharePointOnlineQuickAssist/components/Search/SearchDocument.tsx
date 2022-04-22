import * as React from 'react';
import {  
    PrimaryButton,
    TextField,
    Label,
    ComboBox,
    IComboBox,
    IComboBoxOption,
  } from 'office-ui-fabric-react/lib/index';
import RestAPIHelper from '../../../Helpers/RestAPIHelper';
import { ISharePointOnlineQuickAssistProps } from '../ISharePointOnlineQuickAssistProps';
import SPOQAHelper from '../../../Helpers/SPOQAHelper';
import SPOQASpinner from '../../../Helpers/SPOQASpinner';
import styles from '../SharePointOnlineQuickAssist.module.scss';
import  { ItemType, ModerationStatusHelper} from '../../../Helpers/ModerationStatusHelper';

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
        remedyStepsShowed:false
      };
    private listTitle:string="";
    private resRef= React.createRef<HTMLDivElement>(); 
    private remedyRef = React.createRef<HTMLDivElement>();
    private websNeedFixNoCrawl:string[] =[];
    private redStyle = "color:red";
    private greenStyle = "color:green";
    private remedyStyle = "color:black";
    private remedySteps:any[] =[]; 
    public render():React.ReactElement<ISharePointOnlineQuickAssistProps>
    {
        return (            
            <div id="SearchDocumentContainer">
                <div className={ styles.row }>
                    <div className={ styles.column }>
                        <div id="QuestionsSection">
                            <TextField
                                    label="Affected Site(press enter for loading libraries/lists):"
                                    multiline={false}
                                    onChange={(e)=>{let text:any = e.target; this.setState({affectedSite:text.value,siteIsVaild:false}); this.resRef.current.innerHTML ="";}}
                                    value={this.state.affectedSite}
                                    required={true}
                                    onKeyDown={(e)=>{if(e.keyCode ===13){this.LoadLists();}}}                          
                            /> 
                            {this.state.siteIsVaild? 
                                <div>
                                    <ComboBox
                                    defaultSelectedKey="-1"
                                    label="Please select the affected library/list"
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
                                        }} 
                                    />   
                                {this.state.affectedLibrary.Title!=""? 
                                <div><TextField
                                label="Affected document full URL:"
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
                                {this.state.isChecked && this.state.siteIsVaild && this.state.affectedLibrary.Title!=""?<Label>Diagnose result:</Label>:null}
                                <div style={{marginLeft:20}} id="SearchDocumentCheckResultDiv" ref={this.resRef}></div>
                            </div>                           
                        <div id="CommandButtonsSection">
                            <PrimaryButton
                                text="Check Issues"
                                style={{ display: 'inline', marginTop: '10px' }}
                                onClick={() => {this.state.siteIsVaild? this.CheckSearchDocument():this.LoadLists();}}
                                />
                            {this.state.needRemedy && !this.state.remedyStepsShowed && this.state.siteIsVaild?
                                <PrimaryButton
                                    text="Show Remedy Steps"
                                    style={{ display: 'inline', marginTop: '10px', marginLeft:"10px"}}
                                    onClick={() => {this.ShowRemedySteps();}}
                                />:null}
                        </div>
                        <div id="RemedyStepsDiv" ref={this.remedyRef}>
                                
                        </div>
                    </div>
                </div>  
            </div>
        );
    }
    
    public async LoadLists()
    {       
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
            SPOQAHelper.ShowMessageBar("Error", `Failed to load lists from the site, please make sure the site URL is correct and you have the permssion, detail error is ${err}`);
        }        
    }
    
    public async CheckSearchDocument()
    {
        if(this.state.affectedLibrary.Title == "" ||this.state.affectedLibrary.Title =="-1")
        {
            SPOQAHelper.ShowMessageBar("Error", "Please select the library!");
            return;
        }

        if(!this.state.affectedDocument || this.state.affectedDocument =="" || this.state.affectedDocument.trim()=="")
        {
            SPOQAHelper.ShowMessageBar("Error", "Please provide the affected document full URL!");
            return;
        }

        SPOQAHelper.ResetFormStaus();        
        this.CleanPreviousResult(); 
        let searched:boolean = false;
        SPOQASpinner.Show("Checking document search issue ......");
        this.listTitle = this.state.affectedLibrary.Title;
        try
        {
           var searchDocRes = await RestAPIHelper.SearchDocumentByFullPath(this.props.spHttpClient, this.state.affectedSite, this.state.affectedDocument);
           console.log(searchDocRes);
           if(searchDocRes.RowCount >0)
           {
                SPOQAHelper.ShowMessageBar("Success", `Searched out ${searchDocRes.RowCount} items, looks like the affected document can be searched.`);   
                searched = true;             
           }           
        }
        catch(err)
        {
            SPOQAHelper.ShowMessageBar("Error",`Get exception when try to SearchDocumentByFullPath with error message ${err}`);             
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
                SPOQAHelper.ShowMessageBar("Error",`Get exception when try to check IsWebNoCrawl with error message ${err}`);                 
            }

            // check library no-index
            try
            {
                var resIsListNoCrawl = await RestAPIHelper.IsListNoCrawl(this.props.spHttpClient, this.state.affectedSite, this.listTitle);
                this.setState({isListNoIndex:resIsListNoCrawl});
                var listNoCrawlMsg =  `<span style="${resIsListNoCrawl? this.redStyle:this.greenStyle}" >The nocrawl ${resIsListNoCrawl?"has":"hasn't"} been enabled for the list ${this.state.affectedLibrary.Title}.</span><br/>`;
                this.resRef.current.innerHTML += listNoCrawlMsg;
                if(resIsListNoCrawl)
                {
                    this.remedySteps.push({
                        message:"Dectected nocrawl for the list.",
                        url:`${this.state.affectedSite}/_layouts/15/advsetng.aspx?List={${this.state.affectedLibrary.Id}}`});
                }
            }
            catch(err)
            {
                SPOQAHelper.ShowMessageBar("Error",`Get exception when try to check IsListNoCrawl with error message ${err}`);                
            }

            // check the list form missed issue
            if(!this.state.isLibrary)
            {
                try
                {
                    var resIsListMissedForm = await RestAPIHelper.IsListMissDisplayForm(this.props.spHttpClient, this.state.affectedSite, this.listTitle);
                    // isMissingDisplayForm
                    this.setState({isMissingDisplayForm:resIsListMissedForm});
                    var listMissedDisplayFormMsg =  `<span style="${resIsListMissedForm? this.redStyle:this.greenStyle}" >The dispalyForm ${resIsListMissedForm?"is":"isn't"} missed for the list ${this.state.affectedLibrary.Title}</span><br/>`;
                    this.resRef.current.innerHTML += listMissedDisplayFormMsg;
                    if(resIsListMissedForm)
                    {
                        this.remedySteps.push({message:`Dectected display form is missing, please use the feature "Missing New/Disp/Edit" Forms to fix it`});
                    }
                }
                catch(err)
                {
                    SPOQAHelper.ShowMessageBar("Error",`Get exception when try to check IsListMissDisplayForm with error message ${err}`);                     
                }
            }

            // Check the approve status, OData__ModerationStatus, https://docs.microsoft.com/en-us/dotnet/api/microsoft.sharepoint.spmoderationstatustype?view=sharepoint-server
            var resNotApprovedItems = await ModerationStatusHelper.GetNotApprovedItems(this.props.spHttpClient, this.state.affectedSite, this.state.affectedLibrary.RootFolder, this.state.affectedDocument);
            if(resNotApprovedItems.success)
               {
                    resNotApprovedItems.items.forEach((item)=>{
                        var notApproveMsg = `<a href='${item.url}'>${item.name}</a>'s approve status is ${item.status}.`;
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
                    var docIsDraftVersionMsg =  `<span style="${resIsDraftVersion? this.redStyle:this.greenStyle}" >The document ${this.state.affectedDocument} ${resIsDraftVersion?"is":"isn't"} in draft version.</span><br/>`;
                    this.resRef.current.innerHTML += docIsDraftVersionMsg;
                    if(resIsDraftVersion)
                    {
                        this.remedySteps.push({
                            message:`The object ${this.state.affectedDocument} is in draft version`,
                            url:this.props.rootUrl + this.state.affectedLibrary.RootFolder});    
                    }                
                }
                catch(err)
                {
                    SPOQAHelper.ShowMessageBar("Error",`Get exception when try to check IsDocumentInDraftVersion with error message ${err}`);                
                } 
           }
           else if(resNotApprovedItems.itemType == ItemType.Folder)
           {
                this.resRef.current.innerHTML += `${this.state.affectedDocument} is a folder, skip draft version checking`;
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
    
    // Auto fix will be deprecated
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
                var fixMissingDisplayFormRes = await RestAPIHelper.FixMissDisForm(this.props.spHttpClient, this.state.affectedSite, this.listTitle);
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
        this.resRef.current.innerHTML ="";
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
                    message:`Dectected nocrawl for the site ${currentWebUrl}.`,
                    url:`${currentWebUrl}/_layouts/15/srchvis.aspx`});
            }
            
            var noIndexMsg = `<span style="${webInfo.NoCrawl? this.redStyle:this.greenStyle}" >The nocrawl ${webInfo.NoCrawl?"has":"hasn't"} been enabled for the site ${currentWebUrl}.</span><br/>`;
            this.resRef.current.innerHTML += noIndexMsg;
            currentWebUrl = await RestAPIHelper.GetParentWebUrl(this.props.spHttpClient, currentWebUrl);
            hasParentWeb = currentWebUrl && currentWebUrl!="";
        }       
    }

    public async ShowRemedySteps()
    {    
        this.remedyRef.current.innerHTML+=`<br/><label class="ms-Label" style='${this.remedyStyle};font-size:14px;font-weight:bold'>Remedy Steps:</label><br/>`;
        // Dispaly remedy steps
        this.remedySteps.forEach(step=>{
            var message =step.message;
            if(step.message[step.message.length-1] ==".")
            {
                message = message.substr(0, step.message.length-1);                
            }
            var fixpage = "";
            if(step.url)
            {
                fixpage = ` can be fixed in <a href='${step.url}' target='_blank'>this page</a>`;
            }
            this.remedyRef.current.innerHTML+=`<div style='${this.remedyStyle};margin-left:20px'>${message}${fixpage}.</div>`;
        }); 

        this.setState({remedyStepsShowed:true});   
    }

    private CleanPreviousResult() {
        this.websNeedFixNoCrawl=[];
        this.resRef.current.innerHTML ="";
        this.remedyRef.current.innerHTML =""; 
        this.remedySteps =[];
        this.setState({isChecked:false, needRemedy:false});
    }

}