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

export default class PermissionQA extends React.Component<ISharePointOnlineQuickAssistProps>
{
    public state = {         
        affectedLibrary:{Title:"",Id:"",RootFolder:"", readSecurity:0, writeSecurity:0},
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
    private remedyStyle = "color:black";
    private resRef= React.createRef<HTMLDivElement>();  
    private remedyRef = React.createRef<HTMLDivElement>();
    public render():React.ReactElement<ISharePointOnlineQuickAssistProps>
    {
        return (            
            <div id="SearchDocumentContainer">
                <div className={ styles.row }>
                    <div className={ styles.column }>
                        <div id="QuestionsSection">
                        <TextField
                                label="Affected User:"
                                multiline={false}
                                onChange={(e)=>{let text:any = e.target; this.setState({affectedUser:text.value});}}
                                value={this.state.affectedUser}
                                required={true}                                                
                        />
                            <TextField
                                    label="Affected Site(press enter for loading libraries/lists):"
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
                                                    writeSecurity:option.data.writeSecurity,
                                                    readSecurity:option.data.readSecurity}, 
                                                isChecked:false}); 
                                                this.resRef.current.innerHTML="";
                                                 this.remedyRef.current.innerHTML="";}} 
                                    />
                                    {this.state.affectedLibrary.Title!=""? 
                                        <div><TextField
                                        label="Affected document full URL:"
                                        multiline={false}
                                        onChange={(e)=>{let text:any = e.target; this.setState({affectedDocument:text.value,isChecked:false});}}
                                        value={this.state.affectedDocument}
                                        required={true}                                                
                                        />                                               
                                        <Label>e.g. {this.props.rootUrl}{this.state.affectedLibrary.RootFolder}/xxxx.xxx, if the affected URL is a site, then fill home page full URL</Label>
                                    </div>:null}                                        
                                </div>:null}
                            </div>
                            <div id="OneDriveSyncDiagnoseResult">
                        {this.state.isChecked && this.state.siteIsVaild?<Label>Diagnose result:</Label>:null}
                                <div style={{marginLeft:20}} id="OneDriveSyncDiagnoseResultDiv" ref={this.resRef}>
                                </div>
                        </div>
                        <div id="CommandButtonsSection">
                            <PrimaryButton
                                text="Check Issues"
                                style={{ display: 'inline', marginTop: '10px' }}
                                onClick={() => {this.state.siteIsVaild? this.CheckPermissionQAIssues():this.LoadLists();}}
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
    
    private async LoadLists()
    {       
        try
        {
            var lists:any = await RestAPIHelper.GetLists(this.props.spHttpClient, this.state.affectedSite);
            let listOptions:IComboBoxOption[] = [];

            // list.BaseType ==1 means this list is a library, otherwise this list is a list
            lists.forEach(list => {
                if(list.BaseType ==1) // OneDrive only can sync library, so only libary need to be checked
                {
                    listOptions.push({ 
                        key:list.Title, 
                        text: list.Title,
                        data:{listId:list.Id, rootFolder:list.RootFolder.ServerRelativeUrl, writeSecurity:list.writeSecurity,readSecurity:list.ReadSecurity}});
                }
            });
            this.setState({siteIsVaild:true, siteLibraries:listOptions});
        }
        catch(err)
        {
            SPOQAHelper.ShowMessageBar("Error", `Failed to load lists from the site, please make sure the site URL is correct and you have the permssion, detail error is ${err}`);
        }        
    }
    
    private async CheckPermissionQAIssues()
    {
        if(this.state.affectedLibrary.Title == "" ||this.state.affectedLibrary.Title =="-1")
        {
            SPOQAHelper.ShowMessageBar("Error", "Please select the library!");
            return;
        }

        SPOQAHelper.ResetFormStaus();
        this.setState({isChecked:false, needRemedy:false, remedyStepsShowed:false});         
        this.remedySteps = []; // Clean RemedySteps
        this.resRef.current.innerHTML = ""; // Clean the OneDriveSyncDiagnoseResultDiv
        this.remedyRef.current.innerHTML =""; // Clean the RemedyStepsDiv
        SPOQASpinner.Show(`Checking ${this.state.affectedUser}'s permission issue for URL ${this.state.affectedDocument}......`); 
        
       // check file without check-in version 
       var hasFileWithOutCheckVersion = await RestAPIHelper.HasFileWithOutCheckInVersion(this.props.spHttpClient, this.state.affectedSite, this.state.affectedLibrary.Id);
       var fileWithOutCheckVersionMsg = `There ${hasFileWithOutCheckVersion?"are ":"isn't any " } documents without check-in version in the library.`;
       this.resRef.current.innerHTML += `<span style='${hasFileWithOutCheckVersion? this.redStyle:this.greenStyle}'>${fileWithOutCheckVersionMsg}</span><br/>`;
       if(hasFileWithOutCheckVersion)
       {
            this.remedySteps.push({
                message:fileWithOutCheckVersionMsg,
                url:`${this.state.affectedSite}/_layouts/15/ManageCheckedOutFiles.aspx?List={${this.state.affectedLibrary.Id}}`
            });
       }
       
       // check file existing
       var isFileExisting = await RestAPIHelper.IsDocumentExisting(this.props.spHttpClient, this.state.affectedSite, this.state.affectedDocument.replace(this.props.rootUrl, ""));
       var fileExistingMsg = `The file ${isFileExisting.success? "can":"can't"} be found.`;
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
                    var hasCustomzationMsg = `This page contains customzations ${customzationNames.length >0? customzationNames.join(","):""}`;
                    this.resRef.current.innerHTML += `<span style='${this.redStyle}'>${hasCustomzationMsg}</span><br/>`;
                }
           }
       }

       // check permission on the page/document
       var hasReadPermissionOnDocument =  await RestAPIHelper.HasPermissionOnDocument(this.props.spHttpClient, this.state.affectedSite, this.state.affectedDocument.replace(this.props.rootUrl, ""), this.state.affectedUser, SP.PermissionKind.viewListItems);
       var readPermssionOnDocumentMsg = `The user ${this.state.affectedUser} ${hasReadPermissionOnDocument? "has":"lacks"} read permission on the document`;
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
       var draftVersionMsg = `The document ${isDraftVersion? "is":"is not"} in draft version`;
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
       var securityLevelIssueMsg = `The library ${hasSecurityLevelIssue? "has":"hasn't"} been set to only the author can read/write the item`;
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
        var lockDownMsg = `Limited-access user permission lockdown mode of the site collection has been ${isLockDownEnabled? "enabled":"disabled"}`;
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
        var hasViewPermssionMsg = `The affected user ${hasViewPermission? "has":"doesn't have"} view permission on the library.`;
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

    private async ShowRemedySteps()
    {    
        this.remedyRef.current.innerHTML+=`<br/><label class="ms-Label" style='${this.remedyStyle};font-size:14px;font-weight:bold'>Remedy Steps:</label><br/>`;
        // Dispaly remedy steps
        this.remedySteps.forEach(step=>{
            var message =step.message;
            if(step.message[step.message.length-1] ==".")
            {
                message = message.substr(0, step.message.length-1);                
            }

            this.remedyRef.current.innerHTML+=`<div style='${this.remedyStyle};margin-left:20px'>${message} can be fixed in <a href='${step.url}' target='_blank'>this page</a>.</div>`;
        }); 

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
                    message:`Try to disable 3rd party code by appending ?disable3PCode to page URL and open it in a new tab`,
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
         
        }

        return {hascustomzation:hascustomzation, customzations:customzationsRes};
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