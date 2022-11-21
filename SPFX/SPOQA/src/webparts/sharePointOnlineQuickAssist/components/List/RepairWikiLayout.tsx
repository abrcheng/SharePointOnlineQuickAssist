import * as React from 'react';
import {  
    DefaultButton,
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
import { Text } from '@microsoft/sp-core-library';
import * as strings from 'SharePointOnlineQuickAssistWebPartStrings';

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
                                label={strings.UW_AffectedWikiPage}
                                multiline={false}
                                onChange={(e)=>{let text:any = e.target; this.setState({affectedWiki:text.value});}}
                                value={this.state.affectedWiki}
                                required={true}                        
                        />                         
                        <Label>e.g. https://contoso.sharepoint.com/sites/Intranet/SitePages/Wiki.aspx </Label>   
                        </div>
                    <div id="CommandButtonsSection">
                        <DefaultButton
                            text= {strings.CheckIssues}
                            style={{ display: 'inline', marginTop: '10px' }}
                            onClick={() => {this.resRef.current.innerHTML=""; this.CheckWikiPage();}}
                            />
                    </div>
                     
                        <div id="DiagResultSection">
                        {this.state.siteIsVaild&&this.state.affectedWiki!="" && this.state.isChecked?<Label>{strings.DiagnoseResult}</Label>:null}
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
            SPOQASpinner.Show(Text.format(strings.UW_CheckingWikiLayout,this.state.affectedWiki));
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
                    declaredLayout = strings.UW_WikiLayout_OneColumn;
                    break;
                }
                case "false,false,2":{
                    declaredLayout = strings.UW_WikiLayout_OneColumnWithSideBarOrTwoColumn;
                    break;
                }
                case "true,false,2":{
                    declaredLayout = strings.UW_WikiLayout_TwoColumnWithHeader;
                    break;
                }
                case "true,true,2":{
                    declaredLayout = strings.UW_WikiLayout_TwoColumnWithHeaderAndFooter;
                    break;
                }
                case "false,false,3":{
                    declaredLayout = strings.UW_WikiLayout_ThreeColumn;
                    break;
                }
                case "true,false,3":{
                    declaredLayout = strings.UW_WikiLayout_ThreeColumnWithHeader;
                    break;
                }
                case "true,true,3":{
                    declaredLayout = strings.UW_WikiLayout_ThreeColumnWithHeaderAndFooter;
                    break;
                }
                default:{
                    declaredLayout = strings.UW_WikiLayout_InvalidLayout;
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
                        detectedLayout = strings.UW_WikiLayout_OneColumn;
                    } else if (col==2){
                        detectedLayout = strings.UW_WikiLayout_OneColumnWithSideBarOrTwoColumn;
                    } else if (col==3){
                        detectedLayout = strings.UW_WikiLayout_ThreeColumn;
                    } else {
                        detectedLayout = strings.UW_WikiLayout_InvalidLayout;
                        isInvalidDetectedLayout = true;
                    }
                    break;
                }
                case 2:{
                    col = wikiField.querySelectorAll("#layoutsTable>tbody>tr:nth-child(2)>td").length;
                    if (col==2){
                        detectedLayout = strings.UW_WikiLayout_TwoColumnWithHeader;
                    } else if (col==3){
                        detectedLayout = strings.UW_WikiLayout_ThreeColumnWithHeader;
                    } else {
                        detectedLayout = strings.UW_WikiLayout_InvalidLayout;
                        isInvalidDetectedLayout = true;
                    }
                    break;
                }
                case 3:{
                    col = wikiField.querySelectorAll("#layoutsTable>tbody>tr:nth-child(2)>td").length;
                    if (col==2){
                        detectedLayout = strings.UW_WikiLayout_TwoColumnWithHeaderAndFooter;
                    } else if (col==3){
                        detectedLayout = strings.UW_WikiLayout_ThreeColumnWithHeaderAndFooter;
                    } else {
                        detectedLayout = strings.UW_WikiLayout_InvalidLayout;
                        isInvalidDetectedLayout = true;
                    }
                    break;
                }
                default:{
                    detectedLayout = strings.UW_WikiLayout_InvalidLayout;
                    isInvalidDetectedLayout = true;
                    break;
                }
            }
            this.setState({isInvalidDetectedLayout:isInvalidDetectedLayout});
            this.setState({isChecked:true});

            this.resRef.current.innerHTML += `<div style=color:${isInvalidDeclaredLayout? "Red":"Green"}>${Text.format(strings.UW_DeclaredLayout,declaredLayout)}</div>`;
            this.resRef.current.innerHTML += `<div style=color:${isInvalidDetectedLayout? "Red":"Green"}>${Text.format(strings.UW_DetectedLayout,detectedLayout)}</div>`;
            this.resRef.current.innerHTML += `<div style=color:${declaredLayout!=detectedLayout? "Red":"Green"}>${declaredLayout!=detectedLayout? strings.UW_LayoutNotMatch:strings.UW_LayoutMatch}</div>`;
            if (isInvalidDeclaredLayout || isInvalidDetectedLayout || declaredLayout!=detectedLayout){
                this.resRef.current.innerHTML += `<div style=color:Red>${Text.format(strings.UW_IssueDetected,this.remedyStepWikiURL)}</div>`;
            } else {
                this.resRef.current.innerHTML += `<div style=color:Green>${strings.UW_NoIssueDetected}</div>`;
            }

            SPOQASpinner.Hide();

        }
        catch(err)
        {
            SPOQASpinner.Hide();
            SPOQAHelper.ShowMessageBar("Error", Text.format(strings.UW_Ex_FailedGetPageInfo,err));
        }        
    }    
}