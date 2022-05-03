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
import FormsHelper from '../../../Helpers/FormsHelper';
import { ISharePointOnlineQuickAssistProps } from '../ISharePointOnlineQuickAssistProps';
import SPOQAHelper from '../../../Helpers/SPOQAHelper';
import SPOQASpinner from '../../../Helpers/SPOQASpinner';
import styles from '../SharePointOnlineQuickAssist.module.scss';
import * as strings from 'SharePointOnlineQuickAssistWebPartStrings';
export default class RepairFormQA extends React.Component<ISharePointOnlineQuickAssistProps>
{
    public state = {
        affectedList:"",
        siteLists:[],
        affectedSite:this.props.webAbsoluteUrl,
        isLibrary:false,
        siteIsVaild:false,
        isMissingDisplayForm:false,
        isMissingNewForm:false,
        isMissingEditForm:false,
        isChecked:false
      };
    private listTitle:string="";
    public render():React.ReactElement<ISharePointOnlineQuickAssistProps>
    {
        return (            
            <div id="SearchDocumentContainer">
                <div className={ styles.row }>
                    <div className={ styles.column }>
                        <div id="QuestionsSection">
                            <TextField
                                    label={strings.MF_Label_AffectedSite}
                                    multiline={false}
                                    onChange={(e)=>{let text:any = e.target; this.setState({affectedSite:text.value,siteIsVaild:false});}}
                                    value={this.state.affectedSite}
                                    required={true}
                                    onKeyDown={(e)=>{if(e.keyCode ===13){this.LoadLists();}}}                          
                            /> 
                            {this.state.siteIsVaild? 
                                <div>
                                    <ComboBox
                                    defaultSelectedKey="-1"
                                    label={strings.MF_Label_SelectList}
                                    allowFreeform
                                    autoComplete="on"
                                    options={this.state.siteLists} 
                                    required={true}                    
                                    onChange ={(ev: React.FormEvent<IComboBox>, option?: IComboBoxOption): void => {
                                        this.setState({affectedList: option.key, isLibrary:option.key.toString().endsWith("#1"), isChecked:false});}} 
                                    />     
                                </div>: null}
                            </div>
                            {this.state.siteIsVaild&&this.state.affectedList!="" && this.state.isChecked? 
                                <div id="SearchDocumentCheckResultSection">
                                    <Label>{strings.MF_Label_DiagnoseResult}</Label>
                                    {this.state.isMissingDisplayForm?<Label style={{"color":"Red",marginLeft:"20px"}}>{strings.MF_Message_DispFormMiss} {this.listTitle}</Label>:
                                        <Label style={{"color":"Green",marginLeft:"20px"}}>{strings.MF_Message_DispFormExist}</Label>}
                                    {this.state.isMissingNewForm?<Label style={{"color":"Red",marginLeft:"20px"}}>{strings.MF_Message_NewFormMiss} {this.listTitle}</Label>:
                                        <Label style={{"color":"Green",marginLeft:"20px"}}>{strings.MF_Message_NewFormExist}</Label>}
                                    {this.state.isMissingEditForm?<Label style={{"color":"Red",marginLeft:"20px"}}>{strings.MF_Message_EditFormMiss} {this.listTitle}</Label>:
                                        <Label style={{"color":"Green",marginLeft:"20px"}}>{strings.MF_Message_EditFormExist}</Label>}
                                </div>:null
                            }
                        <div id="CommandButtonsSection">
                            <PrimaryButton
                                text="Check Issues"
                                style={{ display: 'inline', marginTop: '10px' }}
                                onClick={() => {this.state.siteIsVaild? this.CheckListForms():this.LoadLists();}}
                                />
                            {this.state.isChecked && (this.state.isMissingDisplayForm || this.state.isMissingNewForm || this.state.isMissingEditForm)?
                                <PrimaryButton
                                    text="Fix Issues"
                                    style={{ display: 'inline', marginTop: '10px', marginLeft:"10px"}}
                                    onClick={() => {this.FixIssues();}}
                                />:null}
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
                listOptions.push({ key:list.Title+"#"+list.BaseType, text: list.Title});
            });
            this.setState({siteIsVaild:true, siteLists:listOptions});
        }
        catch(err)
        {
            SPOQAHelper.ShowMessageBar("Error", `${strings.MF_Ex_LoadListsError} ${err}`);
        }        
    }
    
    public async CheckListForms()
    {
        if(this.state.affectedList == "" ||this.state.affectedList =="-1")
        {
            SPOQAHelper.ShowMessageBar("Error", `${strings.MF_Ex_ListNotSelected}`);
            return;
        }

        SPOQAHelper.ResetFormStaus();
        this.setState({isChecked:false});
        this.setState({isMissingDisplayForm:false});
        this.setState({isMissingNewForm:false});
        this.setState({isMissingEditForm:false});
        SPOQASpinner.Show(`${strings.MF_Message_CheckingForms}`);
        this.listTitle = this.state.affectedList.substr(0,this.state.affectedList.length-2);
        
        // check the list form missed issue
        try
        {
            // isMissingDisplayForm
            var resIsListMissedForm1 = await RestAPIHelper.IsListMissDisplayForm(this.props.spHttpClient, this.state.affectedSite, this.listTitle);
            this.setState({isMissingDisplayForm:resIsListMissedForm1});
            // isMissingNewForm
            var resIsListMissedForm2 = await RestAPIHelper.IsListMissNewForm(this.props.spHttpClient, this.state.affectedSite, this.listTitle);
            this.setState({isMissingNewForm:resIsListMissedForm2});
            // isMissingEditForm
            var resIsListMissedForm3 = await RestAPIHelper.IsListMissEditForm(this.props.spHttpClient, this.state.affectedSite, this.listTitle);
            this.setState({isMissingEditForm:resIsListMissedForm3});
        }
        catch(err)
        {
            SPOQAHelper.ShowMessageBar("Error",`${strings.MF_Ex_CheckFormsError} ${err}`);                     
        }

        this.setState({isChecked:true});           

        SPOQASpinner.Hide();
    }

    public async FixIssues()
    {
        SPOQAHelper.ResetFormStaus();
        SPOQASpinner.Show(`${strings.MF_Message_FixForms}`);
        let hasError:boolean = false;
        
        if(this.state.isMissingDisplayForm)
        {
            try
            {
                await FormsHelper.FixMissDisForm(this.props.spHttpClient, this.state.affectedSite, this.listTitle);
            }
            catch(err)
            {
                SPOQAHelper.ShowMessageBar("Error",`${strings.MF_Ex_FixDispFormError} ${err}`);
                hasError = true;
            }
        }

        if(this.state.isMissingNewForm)
        {
            try
            {
                await FormsHelper.FixMissNewForm(this.props.spHttpClient, this.state.affectedSite, this.listTitle);
            }
            catch(err)
            {
                SPOQAHelper.ShowMessageBar("Error",`${strings.MF_Ex_FixNewFormError} ${err}`);
                hasError = true;
            }
        }

        if(this.state.isMissingEditForm)
        {
            try
            {
                await FormsHelper.FixMissEditForm(this.props.spHttpClient, this.state.affectedSite, this.listTitle);
            }
            catch(err)
            {
                SPOQAHelper.ShowMessageBar("Error",`${strings.MF_Ex_FixEditFormError} ${err}`);
                hasError = true;
            }
        }

        if(!hasError)
        {
            SPOQAHelper.ShowMessageBar("Success", `${strings.MF_Message_FixedAll}`);
            this.setState({isChecked:false});
            
        }

        SPOQASpinner.Hide();
    }
}