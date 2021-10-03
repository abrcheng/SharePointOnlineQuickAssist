import * as React from 'react';
import {  
    PrimaryButton,
    TextField,
    Label,
    ComboBox,
    IComboBox,
    IComboBoxOption,
  } from 'office-ui-fabric-react/lib/index';
import RestAPIHelper from '../../Helpers/RestAPIHelper';
import { ISharePointOnlineQuickAssistProps } from './ISharePointOnlineQuickAssistProps';
import SPOQAHelper from '../../Helpers/SPOQAHelper';
export default class SearchDocumentQA extends React.Component<ISharePointOnlineQuickAssistProps>
{
    public state = {
        affectedDocument:"",
        affectedLibrary:"",
        siteLibraries:[],         
        affectedSite:this.props.webAbsoluteUrl,
        isLibrary:false,
        siteIsVaild:false,
        isListNoIndex:false,
        isWebNoIndex:false,
        isDraftVersion:false,
        isMissingDisplayForm:false,
        isChecked:false
      };
    private listTitle:string="";
    public render():React.ReactElement<ISharePointOnlineQuickAssistProps>
    {
        return (            
            <div>
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
                        options={this.state.siteLibraries} 
                        required={true}                    
                        onChange ={(ev: React.FormEvent<IComboBox>, option?: IComboBoxOption): void => {
                            this.setState({affectedLibrary: option.key, isLibrary:option.key.toString().endsWith("#1"), isChecked:false});}} 
                        />   
                    {this.state.affectedLibrary!=""? 
                       <div><TextField
                       label="Affected document full URL:"
                       multiline={false}
                       onChange={(e)=>{let text:any = e.target; this.setState({affectedDocument:text.value,isChecked:false});}}
                       value={this.state.affectedDocument}
                       required={true}                                                
                       />
                            {!this.state.isLibrary?<Label>e.g. {this.state.affectedSite}/Lists/{this.state.affectedLibrary.substr(0,this.state.affectedLibrary.length-2)}/DispForm.aspx?ID=xxx</Label>
                                :<Label>e.g. {this.state.affectedSite}/{this.state.affectedLibrary.substr(0,this.state.affectedLibrary.length-2)}/xxxx.xxxx</Label>}
                       </div>:null}                
                    </div>: null}
                    {this.state.siteIsVaild&&this.state.affectedLibrary!="" && this.state.isChecked? 
                        <div id="SearchDocumentCheckResultSection">
                            <Label>Diagnose result,</Label>
                            {this.state.isWebNoIndex?<Label style={{"color":"Red",marginLeft:20}} >The nocrawl has been enabled for the site {this.state.affectedSite}</Label>:
                                <Label style={{"color":"Green",marginLeft:20}}>The nocrawl hasn't been enabled for the site {this.state.affectedSite}</Label>}
                            {this.state.isListNoIndex?<Label style={{"color":"Red",marginLeft:20}}>The nocrawl has been enabled for this list {this.listTitle}</Label>:
                                <Label style={{"color":"Green",marginLeft:20}}>The nocrawl hasn't been enabled for this list {this.listTitle}</Label>}
                            {this.state.isMissingDisplayForm?<Label style={{"color":"Red",marginLeft:20}}>The dispalyForm is missed for the list {this.listTitle}</Label>:
                                <Label style={{"color":"Green",marginLeft:20}}>The dispalyForm is not missed</Label>}
                            {this.state.isDraftVersion?<Label style={{"color":"Red",marginLeft:20}}>The document {this.state.affectedDocument} is in draft version</Label>:
                                <Label style={{"color":"Green",marginLeft:20}}>The document {this.state.affectedDocument} is in major version</Label>}
                        </div>:null
                    }
                  <PrimaryButton
                      text="Check Search Document"
                      style={{ display: 'block', marginTop: '10px' }}
                      onClick={() => {this.CheckSearchDocument();}}
                    />
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
            this.setState({siteIsVaild:true, siteLibraries:listOptions});
        }
        catch(err)
        {
            SPOQAHelper.ShowMessageBar("Error", `Failed to load lists from the site, please make sure the site URL is correct and you have the permssion, detail error is ${err}`);
        }        
    }
    
    public async CheckSearchDocument()
    {
        this.setState({isChecked:false});
        this.listTitle = this.state.affectedLibrary.substr(0,this.state.affectedLibrary.length-2);
        try
        {
           var searchDocRes = await RestAPIHelper.SearchDocumentByFullPath(this.props.spHttpClient, this.state.affectedSite, this.state.affectedDocument);
           console.log(searchDocRes);
           if(searchDocRes.RowCount >0)
           {
                SPOQAHelper.ShowMessageBar("Success", `Searched out ${searchDocRes.RowCount} items, looks like the affected document can be searched.`);
                return;
           }           
        }
        catch(err)
        {
            SPOQAHelper.ShowMessageBar("Error",`Get exception when try to SearchDocumentByFullPath with error message ${err}`);
        }

        // Check web no-index
        try
        {
            var noCrawl = await RestAPIHelper.IsWebNoCrawl(this.props.spHttpClient, this.state.affectedSite);
            this.setState({isWebNoIndex:noCrawl});            
        }
        catch(err)
        {
            SPOQAHelper.ShowMessageBar("Error",`Get exception when try to check IsWebNoCrawl with error message ${err}`);
            return;
        }

        // check library no-index
        try
        {
            var resIsListNoCrawl = await RestAPIHelper.IsListNoCrawl(this.props.spHttpClient, this.state.affectedSite, this.listTitle);
            this.setState({isListNoIndex:resIsListNoCrawl});
        }
        catch(err)
        {
            SPOQAHelper.ShowMessageBar("Error",`Get exception when try to check IsListNoCrawl with error message ${err}`);
            return;
        }

        // check the list form missed issue
        if(!this.state.isLibrary)
        {
            try
            {
                var resIsListMissedForm = await RestAPIHelper.IsListMissDisplayForm(this.props.spHttpClient, this.state.affectedSite, this.listTitle);
                // isMissingDisplayForm
                this.setState({isMissingDisplayForm:resIsListMissedForm});
            }
            catch(err)
            {
                SPOQAHelper.ShowMessageBar("Error",`Get exception when try to check IsListMissDisplayForm with error message ${err}`);
                return;
            }
        }

        // check the draft version issue
        try
        {
            var resIsDraftVersion = await RestAPIHelper.IsDocumentInDraftVersion(this.props.spHttpClient, this.state.affectedSite, this.state.isLibrary, this.listTitle,this.state.affectedDocument);
            this.setState({isDraftVersion:resIsDraftVersion});
        }
        catch(err)
        {
            SPOQAHelper.ShowMessageBar("Error",`Get exception when try to check IsDocumentInDraftVersion with error message ${err}`);
        }

        this.setState({isChecked:true});
    }
}