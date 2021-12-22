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
                 <div id="QuestionsSection">
                    <TextField
                            label="Affected Site(press enter for loading libraries/lists):"
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
                            label="Please select the affected library/list"
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
                            <Label>Diagnose result,</Label>
                            {this.state.isMissingDisplayForm?<Label style={{"color":"Red",marginLeft:"20px"}}>The dispForm is missing for {this.listTitle}</Label>:
                                <Label style={{"color":"Green",marginLeft:"20px"}}>The dispForm exists</Label>}
                            {this.state.isMissingNewForm?<Label style={{"color":"Red",marginLeft:"20px"}}>The newForm is missing for {this.listTitle}</Label>:
                                <Label style={{"color":"Green",marginLeft:"20px"}}>The newForm exists</Label>}
                            {this.state.isMissingEditForm?<Label style={{"color":"Red",marginLeft:"20px"}}>The editForm is missing for {this.listTitle}</Label>:
                                <Label style={{"color":"Green",marginLeft:"20px"}}>The editForm exists</Label>}
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
            SPOQAHelper.ShowMessageBar("Error", `Failed to load lists from the site, please make sure the site URL is correct and you have the permssion, detail error is ${err}`);
        }        
    }
    
    public async CheckListForms()
    {
        if(this.state.affectedList == "" ||this.state.affectedList =="-1")
        {
            SPOQAHelper.ShowMessageBar("Error", "Please select the list/library!");
            return;
        }

        SPOQAHelper.ResetFormStaus();
        this.setState({isChecked:false});
        this.setState({isMissingDisplayForm:false});
        this.setState({isMissingNewForm:false});
        this.setState({isMissingEditForm:false});
        SPOQASpinner.Show("Checking if the forms exist ......");
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
            SPOQAHelper.ShowMessageBar("Error",`Get exception when try to check forms with error message ${err}`);                     
        }

        this.setState({isChecked:true});           

        SPOQASpinner.Hide();
    }

    public async FixIssues()
    {
        SPOQAHelper.ResetFormStaus();
        SPOQASpinner.Show("Repair missing forms ......");
        let hasError:boolean = false;
        
        if(this.state.isMissingDisplayForm)
        {
            try
            {
                await RestAPIHelper.FixMissDisForm(this.props.spHttpClient, this.state.affectedSite, this.listTitle);
            }
            catch(err)
            {
                SPOQAHelper.ShowMessageBar("Error",`Get exception when try fixing the display form with error message ${err}`);
                hasError = true;
            }
        }

        if(this.state.isMissingNewForm)
        {
            try
            {
                await RestAPIHelper.FixMissNewForm(this.props.spHttpClient, this.state.affectedSite, this.listTitle);
            }
            catch(err)
            {
                SPOQAHelper.ShowMessageBar("Error",`Get exception when try fixing the new form with error message ${err}`);
                hasError = true;
            }
        }

        if(this.state.isMissingEditForm)
        {
            try
            {
                await RestAPIHelper.FixMissEditForm(this.props.spHttpClient, this.state.affectedSite, this.listTitle);
            }
            catch(err)
            {
                SPOQAHelper.ShowMessageBar("Error",`Get exception when try fixing the edit form with error message ${err}`);
                hasError = true;
            }
        }

        if(!hasError)
        {
            SPOQAHelper.ShowMessageBar("Success", `Fixed all missing forms`);
            this.setState({isChecked:false});
            
        }

        SPOQASpinner.Hide();
    }
}