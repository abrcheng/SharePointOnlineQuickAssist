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
export default class RepairWikiLayoutQA extends React.Component<ISharePointOnlineQuickAssistProps>
{
    public state = {
        affectedWiki:"",
        pageListId:"",
        pageItemId:"",
        affectedSite:this.props.webAbsoluteUrl,
        siteIsVaild:false,
        isChecked:false,
        detectedLayout:"",
        declaredLayout:"",
        isInvalidDeclaredLayout:false,
        isInvalidDetectedLayout:false
      };
    private resRef= React.createRef<HTMLDivElement>();  
    private remedyStepWikiURL = "https://github.com/abrcheng/SharePointOnlineQuickAssist/blob/main/KBs/UneditableClassicWiki/ReadMe.md";

    public render():React.ReactElement<ISharePointOnlineQuickAssistProps>
    {
        return (    
            <div className={ styles.row }>
                <div className={ styles.column }>
                    <div id="QuestionsSection">
                        <TextField
                                label="Affected Wiki page:"
                                multiline={false}
                                onChange={(e)=>{let text:any = e.target; this.setState({affectedWiki:text.value});}}
                                value={this.state.affectedWiki}
                                required={true}                        
                        />                         
                        </div>
                    <div id="CommandButtonsSection">
                        <PrimaryButton
                            text="Check Issues"
                            style={{ display: 'inline', marginTop: '10px' }}
                            onClick={() => {this.resRef.current.innerHTML=""; this.CheckWikiPage();}}
                            />
                    </div>
                     
                        <div id="DiagResultSection">
                        {this.state.siteIsVaild&&this.state.affectedWiki!="" && this.state.isChecked?<Label>Diagnose result:</Label>:null}
                            <div style={{marginLeft:20}} id="DiagResultDiv" ref={this.resRef}></div>
                        </div>
                    
                </div>
            </div>         
        );
    }
    
    public async CheckWikiPage()
    {       
        try
        {
            SPOQASpinner.Show(`Checking the page layout of ${this.state.affectedWiki}......`);
            this.resRef.current.innerHTML = "";
            var item:Document = await RestAPIHelper.GetPageByUrl(this.props.spHttpClient, this.state.affectedWiki);
            var match = item.body.innerHTML.match(/var _spPageContextInfo=({.*});/);
            var wikiContext = JSON.parse(match[1]);
            this.setState({siteIsVaild:true});
            console.log(wikiContext);

            var wikiField:Document = await RestAPIHelper.GetWikiField(this.props.spHttpClient, wikiContext.webAbsoluteUrl, wikiContext.pageListId, wikiContext.pageItemId);
            console.log(wikiField);
            var layoutsData = wikiField.querySelector("#layoutsData").innerHTML;
            console.log(layoutsData);

            var declaredLayout = "";
            var isInvalidDeclaredLayout = false;
            switch (layoutsData.replace(/\s/gm,'')){
                case "false,false,1":{
                    declaredLayout = "One column.";
                    break;
                }
                case "false,false,2":{
                    declaredLayout = "One column with sitebar or Two column.";
                    break;
                }
                case "true,false,2":{
                    declaredLayout = "Two column with header.";
                    break;
                }
                case "true,true,2":{
                    declaredLayout = "Two column with header and footer.";
                    break;
                }
                case "false,false,3":{
                    declaredLayout = "Three column.";
                    break;
                }
                case "true,false,3":{
                    declaredLayout = "Three column with header.";
                    break;
                }
                case "true,true,3":{
                    declaredLayout = "Three column with header and footer.";
                    break;
                }
                default:{
                    declaredLayout = "Invalid layout.";
                    isInvalidDeclaredLayout = true;
                    break;
                }
            }
            this.setState({isInvalidDeclaredLayout:isInvalidDeclaredLayout});
            
            var detectedLayout = "";
            var isInvalidDetectedLayout = false;
            var row = wikiField.querySelectorAll("#layoutsTable>tbody>tr").length;
            var col = -1;
            switch (row){
                case 1:{
                    col = wikiField.querySelectorAll("#layoutsTable>tbody>tr>td").length;
                    if (col==1){
                        detectedLayout = "One column.";
                    } else if (col==2){
                        detectedLayout = "One column with sitebar or Two column.";
                    } else if (col==3){
                        detectedLayout = "Three column.";
                    } else {
                        detectedLayout = "Invalid layout.";
                    }
                    break;
                }
                case 2:{
                    col = wikiField.querySelectorAll("#layoutsTable>tbody>tr:nth-child(2)>td").length;
                    if (col==2){
                        detectedLayout = "Two column with header.";
                    } else if (col==3){
                        detectedLayout = "Three column with header.";
                    } else {
                        detectedLayout = "Invalid layout.";
                    }
                    break;
                }
                case 3:{
                    col = wikiField.querySelectorAll("#layoutsTable>tbody>tr:nth-child(2)>td").length;
                    if (col==2){
                        detectedLayout = "Two column with header and footer.";
                    } else if (col==3){
                        detectedLayout = "Three column with header and footer.";
                    } else {
                        detectedLayout = "Invalid layout.";
                    }
                    break;
                }
                default:{
                    detectedLayout = "Invalid layout.";
                    isInvalidDetectedLayout = true;
                    break;
                }
            }
            this.setState({isInvalidDetectedLayout:isInvalidDetectedLayout});
            this.setState({isChecked:true});

            this.resRef.current.innerHTML += `<div style=color:${isInvalidDeclaredLayout? "Red":"Green"}>The page layout has been declared as: ${declaredLayout}</div>`;
            this.resRef.current.innerHTML += `<div style=color:${isInvalidDetectedLayout? "Red":"Green"}>The real layout detected from saved HTML elements is: ${detectedLayout}</div>`;
            this.resRef.current.innerHTML += `<div style=color:${declaredLayout!=detectedLayout? "Red":"Green"}>Page layout ${declaredLayout!=detectedLayout? "is not":"is"} matching the declaration</div>`;
            if (isInvalidDeclaredLayout || isInvalidDetectedLayout || declaredLayout!=detectedLayout){
                this.resRef.current.innerHTML += `<div style=color:Red>The diag found the page has layout issue which could cause the ribbon menu being grayed-out and the page uneditable. Please check <a href="${this.remedyStepWikiURL}">this page</a> to fix the issue.</div>`;
            } else {
                this.resRef.current.innerHTML += `<div style=color:Green>The diag didn't find any issue.</div>`;
            }

            SPOQASpinner.Hide();

        }
        catch(err)
        {
            SPOQASpinner.Hide();
            SPOQAHelper.ShowMessageBar("Error", `Failed to get page info, please make sure the page URL is correct and you have the permssion. Detail: ${err}`);
        }        
    }    
}