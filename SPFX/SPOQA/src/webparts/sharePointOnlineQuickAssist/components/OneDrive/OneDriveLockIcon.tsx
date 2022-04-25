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
import {RemedyHelper} from '../../../Helpers/RemedyHelper';

export default class OneDriveLockIconQA extends React.Component<ISharePointOnlineQuickAssistProps>
{
    public state = {         
        affectedLibrary:"",
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
                                        this.setState({affectedLibrary: option.key, isChecked:false}); this.resRef.current.innerHTML=""; this.remedyRef.current.innerHTML="";}} 
                                    />                                      
                                </div>: null}
                            </div>
                            <div id="OneDriveSyncDiagnoseResult">
                                {this.state.isChecked && this.state.siteIsVaild?<Label>Diagnose result:</Label>:null}
                                <div style={{marginLeft:20}} id="OneDriveSyncDiagnoseResultDiv" ref={this.resRef}></div>
                        </div>
                        <div id="CommandButtonsSection">
                            <PrimaryButton
                                text="Check Issues"
                                style={{ display: 'inline', marginTop: '10px' }}
                                onClick={() => {this.state.siteIsVaild? this.CheckOneDriveLockIconQAIssues():this.LoadLists();}}
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
                if(list.BaseType ==1) // OneDrive only can sync library, so only libary need to be checked
                {
                    listOptions.push({ key:list.Title+"#"+list.Id, text: list.Title});
                }
            });
            this.setState({siteIsVaild:true, siteLibraries:listOptions});
        }
        catch(err)
        {
            SPOQAHelper.ShowMessageBar("Error", `Failed to load lists from the site, please make sure the site URL is correct and you have the permssion, detail error is ${err}`);
        }        
    }
    
    public async CheckOneDriveLockIconQAIssues()
    {
        if(this.state.affectedLibrary == "" ||this.state.affectedLibrary =="-1")
        {
            SPOQAHelper.ShowMessageBar("Error", "Please select the library!");
            return;
        }

        SPOQAHelper.ResetFormStaus();
        this.setState({isChecked:false, needRemedy:false, remedyStepsShowed:false});         
        this.remedySteps = []; // Clean RemedySteps
        this.resRef.current.innerHTML = ""; // Clean the OneDriveSyncDiagnoseResultDiv
        this.remedyRef.current.innerHTML =""; // Clean the RemedyStepsDiv
        SPOQASpinner.Show(`Checking OneDrive read only issue for library ${this.state.affectedLibrary}......`);
        this.listTitle = this.state.affectedLibrary.split("#")[0];
        this.listId = this.state.affectedLibrary.split("#")[1];

        // Check list level settings ExcludeFromOfflineClient,ForceCheckout,DraftVersionVisibility,EnableModeration,ValidationFormula,ValidationMessage
        let listProperties:string[]=["ExcludeFromOfflineClient","ForceCheckout","DraftVersionVisibility","EnableModeration","ValidationFormula","ValidationMessage"];        
        var listInfo = await RestAPIHelper.GetListInfo(this.props.spHttpClient, this.state.affectedSite, this.listTitle, listProperties);
        this.CheckListProperties(listInfo);
        this.setState({isChecked:true});

        // Check list schema 
        let listFields = await RestAPIHelper.GetListFields(this.props.spHttpClient, this.state.affectedSite, this.listTitle);
        this.CheckListFields(listFields);
        
        // Check site level settings (ExcludeFromOfflineClient for current web and its parent webs)
        var websInfo = await RestAPIHelper.GetWebExcludeFromOfflineClient(this.props.spHttpClient, this.state.affectedSite);
        this.CheckSiteProperties(websInfo);

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

        // Check affected user's permssion
        var hasEditPermission = await RestAPIHelper.HasPermissionOnList(this.props.spHttpClient, this.state.affectedSite, this.listTitle, this.state.affectedUser, SP.PermissionKind.editListItems);
        var hasEditPermssionMsg = `The affected user ${hasEditPermission? "has":"doesn't have"} edit permssion on the library.`;
        if(!hasEditPermission)
        {
            this.remedySteps.push({
                message:hasEditPermssionMsg,
                url:`${this.state.affectedSite}/_layouts/15/user.aspx?obj={${this.listId}},doclib&List={${this.listId}}`
            });
        }
        this.resRef.current.innerHTML += `<span style='${!hasEditPermission? this.redStyle:this.greenStyle}'>${hasEditPermssionMsg}</span><br/>`;
        if(this.remedySteps.length >0)
        {
            this.setState({needRemedy:true});
        }

        SPOQASpinner.Hide();
    }

    // Check ExcludeFromOfflineClient,ForceCheckout,DraftVersionVisibility,EnableModeration,ValidationFormula,ValidationMessage
    private CheckListProperties(listInfo:any)
    {
        let excludeFromOfflineClientMsg = `Offline Client Availability for the library has been set to ${!listInfo.ExcludeFromOfflineClient}.`;
        if(listInfo.ExcludeFromOfflineClient)
        {
            this.remedySteps.push({
                message:excludeFromOfflineClientMsg,
                url:`${this.state.affectedSite}/_layouts/15/advsetng.aspx?List={${this.listId}}`
            });
           
        }
        this.resRef.current.innerHTML += `<span style='${listInfo.ExcludeFromOfflineClient? this.redStyle:this.greenStyle}'>${excludeFromOfflineClientMsg}</span><br/>`;

        let forceCheckoutMsg = `Require Check Out for the library has been set to ${listInfo.ForceCheckout}.`;
        if(listInfo.ForceCheckout)
        {
            this.remedySteps.push({
                message:forceCheckoutMsg,
                url:`${this.state.affectedSite}/_layouts/15/LstSetng.aspx?List={${this.listId}}`
            });
        }
        this.resRef.current.innerHTML += `<span style='${listInfo.ForceCheckout? this.redStyle:this.greenStyle}'>${forceCheckoutMsg}</span><br/>`;
        
        let draftVisiblitys = ["Any user who can read items", "Only users who can edit items", "Only users who can approve items (and the author of the item)"];
        let draftVersionVisibilityMsg =`Draft Item Security of this library has been set to ${draftVisiblitys[listInfo.DraftVersionVisibility]}.`;
        if(listInfo.DraftVersionVisibility ==2)
        {
            this.remedySteps.push({
                message:draftVersionVisibilityMsg,
                url:`${this.state.affectedSite}/_layouts/15/LstSetng.aspx?List={${this.listId}}`
            }); 
        }        
        this.resRef.current.innerHTML += `<span style='${listInfo.DraftVersionVisibility==2? this.redStyle:this.greenStyle}'>${draftVersionVisibilityMsg}</span><br/>`;

        let enableModerationMsg = `Content Approval of this library has been set to ${listInfo.EnableModeration}.`;
        if(listInfo.EnableModeration)
        {
            this.remedySteps.push({
                message:enableModerationMsg,
                url:`${this.state.affectedSite}/_layouts/15/LstSetng.aspx?List={${this.listId}}`
            }); 
        }
        this.resRef.current.innerHTML += `<span style='${listInfo.EnableModeration? this.redStyle:this.greenStyle}'>${enableModerationMsg}</span><br/>`;
        
        let validationFormulaMsg=`Validation formula/message of this library is ${listInfo.ValidationFormula?listInfo.ValidationFormula:"null"}/${listInfo.ValidationMessage?listInfo.ValidationMessage:"null"}.`;
        if(listInfo.ValidationFormula || listInfo.ValidationMessage)
        {
            this.remedySteps.push({
                message:validationFormulaMsg,
                url:`${this.state.affectedSite}/_layouts/15/VldSetng.aspx?List={${this.listId}}`
            }); 
        }

        this.resRef.current.innerHTML += `<span style='${listInfo.ValidationFormula || listInfo.ValidationMessage? this.redStyle:this.greenStyle}'>${validationFormulaMsg}</span><br/>`;
    }

    // ExcludeFromOfflineClient
    private CheckSiteProperties(websInfo:any[])
    {
        websInfo.forEach(webInfo =>{
            var webExcludeFromOfflineClientMsg = `Offline Client Availability of the web ${webInfo.webUrl} has been set to ${!webInfo.ExcludeFromOfflineClient}.`;
            if(webInfo.ExcludeFromOfflineClient)
            {
                this.remedySteps.push({
                    message:webExcludeFromOfflineClientMsg,
                    url:webInfo.RemedyUrl
                }); 
            }
            this.resRef.current.innerHTML += `<span style='${webInfo.ExcludeFromOfflineClient? this.redStyle:this.greenStyle}'>${webExcludeFromOfflineClientMsg}</span><br/>`;
        });
    }
    
    // Check Required, ValidationFormula and ValidationMessage for list Fields
    private CheckListFields(fields:any[])
    {   
        let schemaCheckPassed:boolean = true;
        let schemaCheckPassedMsg = "Schema check for this library passed.";
        let re = /\-/gi;
        fields.forEach(field =>{
            if(field.InternalName!="FileLeafRef")  // skip the system field FileLeafRef
            {
                var remdyUrl = `${this.state.affectedSite}/_layouts/15/FldEdit.aspx?List=%7B${this.listId.toUpperCase().replace(re, "%2D")}%7D&Field=${field.InternalName}`;
                var commonMsg = `of the column named ${field.Title} has been set to`;
                if(field.Required)
                {
                    var requiredMsg =`The Required ${commonMsg} true.`;
                    schemaCheckPassed = false;
                    this.remedySteps.push({
                        message:requiredMsg,
                        url:remdyUrl
                    }); 

                    this.resRef.current.innerHTML += `<span style='${this.redStyle}'>${requiredMsg}</span><br/>`;
                }

                if(field.ValidationFormula)
                {
                    schemaCheckPassed = false;
                    var validationFormulaMsg = `The ValidationFormula ${commonMsg} ${field.ValidationFormula}.`;
                    this.remedySteps.push({
                        message:validationFormulaMsg,
                        url:remdyUrl
                    }); 
                    this.resRef.current.innerHTML += `<span style='${this.redStyle}'>${validationFormulaMsg}</span><br/>`;
                }
                if(field.ValidationMessage)
                {
                    schemaCheckPassed = false;
                    var validationMessageMsg = `The ValidationMessage ${commonMsg} ${field.ValidationMessage}.`;
                    this.remedySteps.push({
                        message:validationMessageMsg,
                        url:remdyUrl
                    }); 
                    this.resRef.current.innerHTML += `<span style='${this.redStyle}'>${validationMessageMsg}</span><br/>`;
                }
            }
        });

        if(schemaCheckPassed)
        {
            this.resRef.current.innerHTML += `<span style='${this.greenStyle}'>${schemaCheckPassedMsg}</span><br/>`;
        }
    }

    private async ShowRemedySteps()
    {    
        this.remedyRef.current.innerHTML = RemedyHelper.ShowRemedySteps(this.remedySteps);
        this.setState({remedyStepsShowed:true});   
    }
}