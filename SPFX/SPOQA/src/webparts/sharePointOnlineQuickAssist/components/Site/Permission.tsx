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
import {SPHttpClient} from '@microsoft/sp-http';
import styles from '../SharePointOnlineQuickAssist.module.scss';
import {RemedyHelper} from '../../../Helpers/RemedyHelper';
import  { ItemType, ModerationStatusHelper} from '../../../Helpers/ModerationStatusHelper';
import * as strings from 'SharePointOnlineQuickAssistWebPartStrings';

export default class PermissionQA extends React.Component<ISharePointOnlineQuickAssistProps>
{
    public state = {         
        affectedLibrary:{Title:"",Id:"",RootFolder:"", readSecurity:0, writeSecurity:0,baseType:-1},
        affectedDocument:"",
        siteLibraries:[],         
        affectedSite:this.props.webAbsoluteUrl,        
        siteIsVaild:false,       
        isChecked:false,
        needRemedy:false,
        affectedUser:this.props.currentUser.loginName,
        remedyStepsShowed:false
      };

    private listTitle:string="";
    private listId:string="";
    private remedySteps =[]; 
    private redStyle = "color:red";
    private greenStyle = "color:green";   
    private resRef= React.createRef<HTMLDivElement>();  
    private remedyRef = React.createRef<HTMLDivElement>();
    public render():React.ReactElement<ISharePointOnlineQuickAssistProps>
    {
        return (            
            <div id="PermssionQAContainer">
                <div className={ styles.row }>
                    <div className={ styles.column }>
                        <div id="QuestionsSection">
                        <TextField
                                label={strings.AffectedUser}
                                multiline={false}
                                onChange={(e)=>{let text:any = e.target; this.setState({affectedUser:text.value});}}
                                value={this.state.affectedUser}
                                required={true}                                                
                        />
                            <TextField
                                    label={strings.AffectedSiteLoadList}
                                    multiline={false}
                                    onChange={(e)=>{let text:any = e.target; this.setState({affectedSite:text.value,siteIsVaild:false,isChecked:false}); this.resRef.current.innerHTML=""; this.remedyRef.current.innerHTML="";}}
                                    value={this.state.affectedSite}
                                    required={true}
                                    onKeyDown={(e)=>{if(e.keyCode ===13){this.LoadLists();}}}                          
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
                                                    writeSecurity:option.data.writeSecurity,
                                                    readSecurity:option.data.readSecurity,
                                                    baseType:option.data.baseType}, 
                                                isChecked:false}); 
                                                this.resRef.current.innerHTML="";
                                                 this.remedyRef.current.innerHTML="";}} 
                                    />
                                    {this.state.affectedLibrary.Title!=""? 
                                        <div><TextField
                                        label={strings.AffectedDocument}
                                        multiline={false}
                                        onChange={(e)=>{let text:any = e.target; this.setState({affectedDocument:text.value,isChecked:false});}}
                                        value={this.state.affectedDocument}
                                        required={true}                                                
                                        />                                               
                                        {!(this.state.affectedLibrary.baseType==1)?<Label>e.g. {this.props.rootUrl}{this.state.affectedLibrary.RootFolder}/DispForm.aspx?ID=xxx</Label>
                                            :<Label>e.g. {this.props.rootUrl}{this.state.affectedLibrary.RootFolder}/xxxx.xxx</Label>}
                                    </div>:null}                                        
                                </div>:null}
                            </div>
                    <div id="PermssionDiagnoseResult">
                        {this.state.isChecked && this.state.siteIsVaild?<Label>Diagnose result:</Label>:null}
                                <div style={{marginLeft:20}} id="PermssionDiagnoseResultDiv" ref={this.resRef}>
                                </div>
                        </div>
                        <div id="CommandButtonsSection">
                            <PrimaryButton
                                text={strings.CheckIssues}
                                style={{ display: 'inline', marginTop: '10px' }}
                                onClick={() => {this.state.siteIsVaild? this.CheckPermissionQAIssues():this.LoadLists();}}
                                />
                            {this.state.needRemedy && !this.state.remedyStepsShowed && this.state.siteIsVaild?
                                <PrimaryButton
                                    text={strings.ShowRemedySteps}
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
    
    private async LoadLists()
    {       
        try
        {
            var lists:any = await RestAPIHelper.GetLists(this.props.spHttpClient, this.state.affectedSite);
            let listOptions:IComboBoxOption[] = [];

            // list.BaseType ==1 means this list is a library, otherwise this list is a list
            lists.forEach(list => {
                // if(list.BaseType ==1) // 
                //{
                    listOptions.push({ 
                        key:list.Title, 
                        text: list.Title,
                        data:{listId:list.Id, rootFolder:list.RootFolder.ServerRelativeUrl, writeSecurity:list.writeSecurity,readSecurity:list.ReadSecurity, baseType:list.BaseType}});
                //}
            });
            this.setState({siteIsVaild:true, siteLibraries:listOptions});
        }
        catch(err)
        {
            SPOQAHelper.ShowMessageBar("Error", `${strings.FailedLoadSiteList} ${err}`);
        }        
    }
    
    private async CheckPermissionQAIssues()
    {
        if(this.state.affectedLibrary.Title == "" ||this.state.affectedLibrary.Title =="-1")
        {
            SPOQAHelper.ShowMessageBar("Error", strings.PleaseSelectList);
            return;
        }

        SPOQAHelper.ResetFormStaus();
        this.setState({isChecked:false, needRemedy:false, remedyStepsShowed:false});         
        this.remedySteps = []; // Clean RemedySteps
        this.resRef.current.innerHTML = ""; // Clean the OneDriveSyncDiagnoseResultDiv
        this.remedyRef.current.innerHTML =""; // Clean the RemedyStepsDiv
        SPOQASpinner.Show(`${strings.Checking} ......`); 
        
       // check file without check-in version 
       var hasFileWithOutCheckVersion = await RestAPIHelper.HasFileWithOutCheckInVersion(this.props.spHttpClient, this.state.affectedSite, this.state.affectedLibrary.Id);
       var fileWithOutCheckVersionMsg = hasFileWithOutCheckVersion? strings.PC_DocumentsWithoutCheckin:strings.PC_NoDocumentsWithoutCheckin;
       this.resRef.current.innerHTML += `<span style='${hasFileWithOutCheckVersion? this.redStyle:this.greenStyle}'>${fileWithOutCheckVersionMsg}</span><br/>`;
       if(hasFileWithOutCheckVersion)
       {
            this.remedySteps.push({
                message:fileWithOutCheckVersionMsg,
                url:`${this.state.affectedSite}/_layouts/15/ManageCheckedOutFiles.aspx?List={${this.state.affectedLibrary.Id}}`
            });
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
       
       // check file existing
       var isFileExisting = await RestAPIHelper.IsDocumentExisting(this.props.spHttpClient, this.state.affectedSite, this.state.affectedDocument.replace(this.props.rootUrl, ""), this.state.affectedLibrary.RootFolder, resNotApprovedItems.itemType);
       var fileExistingMsg = isFileExisting.success? strings.PC_FileExistingMsg:strings.PC_FileNotExistingMsg;
       this.resRef.current.innerHTML += `<span style='${!isFileExisting.success? this.redStyle:this.greenStyle}'>${fileExistingMsg}</span><br/>`;
       if(!isFileExisting.success)
       {
            this.remedySteps.push({
                message:fileExistingMsg,
                url:`${this.state.affectedSite}/${this.state.affectedLibrary.RootFolder}`
            });
       }

       if(isFileExisting.success)
       {
           // check customizations (modern/classic) in pages
           if(this.state.affectedDocument.endsWith(".aspx")) 
           {
                var customzationRes =  await (this.CheckCustomziation());
                if(customzationRes.hascustomzation)
                {
                    var customzationNames:string[]=[];
                    customzationRes.customzations.forEach(customzation => {
                        if(customzation.alias!="")
                        {
                            customzationNames.push(customzation.alias);
                        }
                    });
                    var hasCustomzationMsg = `${strings.PC_PageCustomized} ${customzationNames.length >0? customzationNames.join(","):""}`;
                    this.resRef.current.innerHTML += `<span style='${this.redStyle}'>${hasCustomzationMsg}</span><br/>`;
                }
           }
       }

       // check permission on the page/document
       var hasReadPermissionOnDocument =  await RestAPIHelper.HasPermissionOnItem(this.props.spHttpClient, this.state.affectedSite, this.state.affectedDocument.replace(this.props.rootUrl, ""), this.state.affectedUser, SP.PermissionKind.viewListItems,this.state.affectedLibrary.RootFolder, resNotApprovedItems.itemType);
       var readPermssionOnDocumentMsg = hasReadPermissionOnDocument? strings.PC_UserHasPermssionOnDocument:strings.PC_UserHasNoPermssionOnDocument;
       this.resRef.current.innerHTML += `<span style='${!hasReadPermissionOnDocument? this.redStyle:this.greenStyle}'>${readPermssionOnDocumentMsg}</span><br/>`;
       if(!hasReadPermissionOnDocument)
       {
            this.remedySteps.push({
                message:readPermssionOnDocumentMsg,
                url:`${this.state.affectedSite}/${this.state.affectedLibrary.RootFolder}`
            });
       } 

       // check file without major version      
       var isDraftVersion = await RestAPIHelper.IsDocumentInDraftVersion(this.props.spHttpClient, this.state.affectedSite, true, this.state.affectedLibrary.Title,this.state.affectedDocument);
       var draftVersionMsg = isDraftVersion? strings.PC_DocumentIsInDraft:strings.PC_DocumentIsNotInDraft;
       this.resRef.current.innerHTML += `<span style='${isDraftVersion? this.redStyle:this.greenStyle}'>${draftVersionMsg}</span><br/>`;
       if(isDraftVersion)
       {
            this.remedySteps.push({
                message:draftVersionMsg,
                url:`${this.props.rootUrl}/${this.state.affectedLibrary.RootFolder}`
            });
       }

       // check library's read/write security is 2 (only the author can read/write the item) 
       var hasSecurityLevelIssue = this.state.affectedLibrary.readSecurity ===2 || this.state.affectedLibrary.writeSecurity ===2;
       var securityLevelIssueMsg = hasSecurityLevelIssue? strings.PC_ListSecurityLevelHasIssue:strings.PC_ListSecurityLevelHasNoIssue;
       this.resRef.current.innerHTML += `<span style='${hasSecurityLevelIssue? this.redStyle:this.greenStyle}'>${securityLevelIssueMsg}</span><br/>`;
       if(hasSecurityLevelIssue)
       {
            this.remedySteps.push({
                message:securityLevelIssueMsg,
                url:`https://github.com/abrcheng/SharePointOnlineQuickAssist/blob/main/KBs/List/UpdateListReadWriteSecurity.ps1`
            });
       }

       // check lock down mode
       // Check site features, ViewFormPagesLockDown feature ID:7c637b23-06c4-472d-9a9a-7c175762c5c4, Limited-access user permission lockdown mode
        var isLockDownEnabled = await RestAPIHelper.IsSiteFeatureEnabled(this.props.spHttpClient, this.state.affectedSite, "7c637b23-06c4-472d-9a9a-7c175762c5c4");
        var lockDownMsg = isLockDownEnabled? strings.PC_LockDownEnabled:strings.PC_LockDownNotEnabled;
        if(isLockDownEnabled)
        {
            this.remedySteps.push({
                message:lockDownMsg,
                url:`${this.state.affectedSite}/_layouts/15/ManageFeatures.aspx?Scope=Site`
            });
        }
        this.resRef.current.innerHTML += `<span style='${isLockDownEnabled? this.redStyle:this.greenStyle}'>${lockDownMsg}</span><br/>`;

        // Check affected user's permssion on the library/list
        var hasViewPermission = await RestAPIHelper.HasPermissionOnList(this.props.spHttpClient, this.state.affectedSite, this.state.affectedLibrary.Title, this.state.affectedUser, SP.PermissionKind.viewListItems);
        var hasViewPermssionMsg = hasViewPermission? strings.PC_HasViewPermissionOnList:strings.PC_HasNoViewPermissionOnList;
        if(!hasViewPermission)
        {
            this.remedySteps.push({
                message:hasViewPermssionMsg,
                url:`${this.state.affectedSite}/_layouts/15/user.aspx?obj={${this.listId}},doclib&List={${this.listId}}`
            });
        }

        this.resRef.current.innerHTML += `<span style='${!hasViewPermission? this.redStyle:this.greenStyle}'>${hasViewPermssionMsg}</span><br/>`;
        if(this.remedySteps.length >0)
        {
            this.setState({needRemedy:true});
        }

        SPOQASpinner.Hide();
    } 

    private ShowRemedySteps()
    {    
        this.remedyRef.current.innerHTML = RemedyHelper.GetRemedyHtml(this.remedySteps);
        this.setState({remedyStepsShowed:true});   
    }
    
    private async CheckCustomziation()
    {
        var apiUrl = `${this.state.affectedDocument}?asjson=1`;
        var response = await this.props.spHttpClient.get(apiUrl, SPHttpClient.configurations.v1);
        var customzationsRes:any[] = [];
        var isModernPage = false;
        if(response.ok)
        {
            var html = await response.text();
            try
            {
                var modernPage = JSON.parse(html);
                var customzations = modernPage.manifests.filter(m => !m.isInternal);
                customzations.forEach(customzation => {
                    customzationsRes.push(this.GetCustomzationFromManifest(customzation));
                }); 
                isModernPage = true;               
            }
            catch(e) // classic page
            {
                const parser = new DOMParser();
                const parsedDocument = parser.parseFromString(html, "text/html");               
                
                //parsedDocument.querySelectorAll("script")[0].attributes["src"]
                var scripts:NodeListOf<HTMLScriptElement> = parsedDocument.querySelectorAll("script");
                var classicRes:string[]=[];
                for(var index=0; index <scripts.length; index++)
                {                    
                    if(scripts[index].attributes["src"] && scripts[index].src.indexOf(this.props.rootUrl) !=-1)
                    {
                        var resPath = scripts[index].src;
                        if(resPath.indexOf("/_layouts/")=== -1 && (resPath.endsWith(".css")||resPath.endsWith(".js")))
                        {
                            classicRes.push(resPath);                           
                        }
                    }                  
                }

                // parsedDocument.querySelectorAll("link")[0].href
                var links:NodeListOf<HTMLLinkElement> = parsedDocument.querySelectorAll("link");
                for(var indexl=0; indexl < links.length; indexl++)
                {
                    if(links[indexl].attributes["href"] && links[indexl].href.indexOf(this.props.rootUrl) !=-1)
                    {
                        var hrefPath = links[indexl].href;
                        if(hrefPath.indexOf("/_layouts/")=== -1 && (hrefPath.endsWith(".css")||hrefPath.endsWith(".js")))
                        {
                            classicRes.push(hrefPath);                           
                        }
                    }
                }

                // Get SPFx context in classic page JSON.parse(html.substring(129132 +"spClientSidePageContext=".length ,222623+1))
                const spfxContextStartStr = "spClientSidePageContext=";
                var spfxContextStartIndex = html.indexOf(spfxContextStartStr) + spfxContextStartStr.length;
                if(spfxContextStartIndex > spfxContextStartStr.length)
                {
                    var spfxContextEndIndex = html.indexOf("};",spfxContextStartIndex) +1;
                    var spfxContext = JSON.parse(html.substring(spfxContextStartIndex, spfxContextEndIndex));
                    spfxContext.manifests.forEach(manifest => {
                        // componentType:"WebPart", "Extension"
                        if(manifest.componentType === "WebPart" || manifest.componentType === "Extension")
                        {
                            customzationsRes.push(this.GetCustomzationFromManifest(manifest));
                        }
                    });
                }

                if(classicRes.length >=1)
                {
                    customzationsRes.push(
                        {
                            alias:"", // alias is empty means classic page
                            type:"",
                            resPaths:classicRes
                        }
                    );
                }
            }    
        }
        else
        {
            console.log(`Failed to load ${apiUrl}`);
            return {hascustomzation:false};
        } 
        var hascustomzation = false;
        if(customzationsRes.length >0)
        {
            hascustomzation = true;
            if(isModernPage) // modern page
            {
                var disable3rdcodeUrl = `${this.state.affectedDocument}?disable3PCode`;
                this.remedySteps.push({
                    message:strings.PC_TryDisable3PCode,
                    url:`${disable3rdcodeUrl}`
                });
            }
            
            customzationsRes.forEach(customzation=>{
                customzation.resPaths.forEach(resUrl =>{
                    this.remedySteps.push({
                        message:`${resUrl}`,
                        url:`${resUrl.substring(0,resUrl.lastIndexOf("/"))}`
                    });
                });
            });
            
           await this.CheckCustomzationPermission(customzationsRes);
        }
        
        return {hascustomzation:hascustomzation, customzations:customzationsRes};
    }

    private async CheckCustomzationPermission(customzationsRes:any[]) {
        // https://chengc.sharepoint.com/sites/SPOQA/ClientSideAssets/029e4fc2-a440-4f5f-a358-b34a0eca54b5/_api/
        // <service xmlns="http://www.w3.org/2007/app" xmlns:atom="http://www.w3.org/2005/Atom" xml:base="https://chengc.sharepoint.com/sites/SPOQA/_api/">
      
        for(var index =0; index < customzationsRes.length;index++)
        {
            var currentCustomzation = customzationsRes[index];
            for(var indexRes=0; indexRes< currentCustomzation.resPaths.length;indexRes++)
            {
                var resUrl:string = currentCustomzation.resPaths[indexRes];
                var res = await this.props.spHttpClient.get(resUrl, SPHttpClient.configurations.v1);
                var resName = resUrl.substring(resUrl.lastIndexOf("/")+1);
                if(res.ok)
                {
                    res = await this.props.spHttpClient.get(resUrl+"/_api/", SPHttpClient.configurations.v1);
                    if(res.ok)
                    {
                        var resJson = await res.json();

                        // https://chengc.sharepoint.com/sites/SPOQA
                        var resWebUrl = resJson["@odata.context"].replace("/_api/$metadata","");
                        var hasPermission = await RestAPIHelper.HasPermissionOnDocument(this.props.spHttpClient, resWebUrl,resUrl.replace(this.props.rootUrl,""), this.state.affectedUser, SP.PermissionKind.viewListItems);
                        if(!hasPermission)
                        {
                            this.resRef.current.innerHTML += `<span style='${this.redStyle}'>${strings.PC_LackPermissionOn} <a href="${resUrl}"> ${resName}</a>!</span><br/>`;
                        }    
                    }
                    else
                    {
                        this.resRef.current.innerHTML += `<span style='${this.redStyle}'><a href="${resUrl}"> ${resName}</a> ${strings.PC_CanNotLoad}</span><br/>`;
                    }                
                }
                else
                {                   
                    this.resRef.current.innerHTML += `<span style='${this.redStyle}'><a href="${resUrl}"> ${resName}</a> ${strings.PC_CanNotLoad}</span><br/>`;                     
                }
            }
        }        
    }
    
    private GetCustomzationFromManifest(manifest:any)
    {
        // alias, type, resUrls
        var baseUrl = manifest.loaderConfig.internalModuleBaseUrls[0].replace("publiccdn.sharepointonline.com/","").replace("privatecdn.sharepointonline.com/","");
        var resUrls:string[]=[];
        for(var key in manifest.loaderConfig.scriptResources)
        {
           var res = manifest.loaderConfig.scriptResources[key];
           if(res.type === "path" || res.type === "localizedPath")
           {
               // path, paths, defaultPath
               if(res.path)
               {
                    resUrls.push(`${baseUrl}/${res.path}`);
               }

               if(res.defaultPath)
               {
                    resUrls.push(`${baseUrl}/${res.defaultPath}`);
               }

               if(res.paths)
               {
                    for(var pathKey in res.paths)
                    {
                        resUrls.push(`${baseUrl}/${res.paths[pathKey]}`);
                    }
               }
           }
        }

        return {
            alias:manifest.alias,
            type:manifest.type,
            resPaths:resUrls
        };
    }
}