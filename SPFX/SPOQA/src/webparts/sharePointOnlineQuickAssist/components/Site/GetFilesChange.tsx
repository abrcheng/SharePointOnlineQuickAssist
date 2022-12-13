import * as React from 'react'; 
import styles from '../SharePointOnlineQuickAssist.module.scss';
import {  
    DefaultButton,
    MessageBar,
    MessageBarType,
    DatePicker,
    TextField
  } from 'office-ui-fabric-react/lib/index';
  import { ISharePointOnlineQuickAssistProps } from '../ISharePointOnlineQuickAssistProps';
import GraphAPIHelper from '../../../Helpers/GraphAPIHelper';  
import RestAPIHelper from '../../../Helpers/RestAPIHelper';  
import SPOQASpinner from '../../../Helpers/SPOQASpinner';
import SPOQAHelper from '../../../Helpers/SPOQAHelper';
import IFilesGrid from "./IFilesGrid";
import { IFile } from "./IFile";
import * as strings from 'SharePointOnlineQuickAssistWebPartStrings';
export default class GetFilesChange extends React.Component<ISharePointOnlineQuickAssistProps>
{
    private modifiedFiles:IFile[];    
    public state = {
        querySite:this.props.webAbsoluteUrl,
        message:"",
        messageType:MessageBarType.success,
        queried:false,
        queryStartDate:null,
        queryEndDate:null,
        modifiedByUser: this.props.currentUser.loginName,
        pathFilter:""
    };

    public render():React.ReactElement<ISharePointOnlineQuickAssistProps>
    {
        return (
            <div>
                <div className={ styles.row }>
                    <div className={ styles.column }>
                        <div id="IFiles_FilterSection" className={styles.msgrid}>
                            <div className={styles.msrow} id="siteURL_row">
                                <TextField
                                    label={strings.FC_Label_QuerySite}
                                    multiline={false}
                                    onChange={(e)=>{let text:any = e.target; this.setState({querySite:text.value});}}
                                    value={this.state.querySite}
                                    required={true}                        
                                /> 
                            </div>
                            <div className={styles.msrow} id="queryFilter_row">
                                <div className={styles.mscol6}>
                                    <TextField
                                        label={strings.FC_Lable_ModifedUser}                  
                                        multiline={false}
                                        onChange={(e)=>{let text:any = e.target; this.setState({modifiedByUser:text.value});}}
                                        value={this.state.modifiedByUser}
                                    />
                                    </div>          
                                    <div className={styles.mscol6}>
                                    <TextField 
                                        label={strings.FC_Label_PathFilter}
                                        className='ms-Grid-col ms-u-sm6 block'
                                        multiline={false}
                                        onChange={(e)=>{let text:any = e.target; this.setState({pathFilter:text.value});}}
                                        value={this.state.pathFilter}                          
                                    />     
                                </div>
                            </div>
                            <div className={styles.msrow} id="queryDate_row">
                                <div className={styles.mscol6}>
                                    <DatePicker
                                        allowTextInput={true}
                                        isMonthPickerVisible ={false}    
                                        label={strings.FC_Label_StartDate}
                                        placeholder={strings.FC_Message_SelectDate}
                                        ariaLabel="Select a date"
                                        onSelectDate={(e)=>{ this.setState({queryStartDate:e});}}
                                        value={this.state.queryStartDate}                    
                                    />
                                </div>
                                <div className={styles.mscol6}>
                                    <DatePicker
                                        allowTextInput={true}
                                        isMonthPickerVisible ={false}    
                                        label={strings.FC_Label_EndDate}
                                        placeholder={strings.FC_Message_SelectDate}
                                        ariaLabel="Select a date"
                                        onSelectDate={(e)=>{ this.setState({queryEndDate:e});}}
                                        value={this.state.queryEndDate}                           
                                    />
                                </div>
                            </div> 
                        </div>
                    </div>
                </div>
                <div className={ styles.row }>
                    <div className={ styles.column }>
                        <div id="IFiles_CommandButtonsSection">
                            <DefaultButton
                                text={strings.FC_Label_GetFiles}
                                style={{ display: 'inline', marginTop: '10px' }}
                                onClick={() => {SPOQAHelper.ResetFormStaus();this.QueryFiles();}}
                            />
                            <DefaultButton
                                text={strings.FC_Label_Export}
                                style={{ display: 'inline', marginTop: '10px', marginLeft:"10px"}}
                                onClick={() => {this.DoExport();}}
                            />
                        </div>
                    </div>
                </div>
                <div id="IFiles_QueryResultSection">
                    {this.state.queried? <MessageBar id="IFilesMessageBar" messageBarType={this.state.messageType} isMultiline={true}>
                        {this.state.message}
                    </MessageBar>:null}
                    {this.state.queried && this.modifiedFiles.length >0? <IFilesGrid items={this.modifiedFiles}/>:null}
                </div>
            </div>
        );
    }

    private async QueryFiles()
    {
        if(this.state.querySite =="")
    {
      SPOQAHelper.ShowMessageBar("Error", strings.UI_NonAffectedSite);          
      return;
    }
        this.setState({queried:false});
        //Get Site ID
        var siteID = await RestAPIHelper.GetSiteId(this.props.spHttpClient, this.state.querySite);
        if(siteID)
        {   
            this.modifiedFiles = [];
            SPOQASpinner.Show(`${strings.FC_Message_Quering}`);
            
            var drives = await RestAPIHelper.GetDrives(this.props.spHttpClient,this.state.querySite);
            console.log(drives);

            for(var k=0; k<drives.length; k++)
            {
                console.log(drives[k]['id']);
                
                //Get files
                /*
                @odata.type: "#microsoft.graph.driveItem"
                cTag: "\"c:{94251CD1-4A09-4F49-A4AD-DDDA14654439},8\""
                createdBy: {user: {…}}
                createdDateTime: "2022-02-07T03:38:00Z"
                eTag: "\"{94251CD1-4A09-4F49-A4AD-DDDA14654439},7\""
                file: {mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', hashes: {…}}
                fileSystemInfo: {createdDateTime: '2022-02-07T03:38:00Z', lastModifiedDateTime: '2022-03-10T09:43:52Z'}
                id: "01JZT46KORDQSZICKKJFH2JLO53IKGKRBZ"
                lastModifiedBy:
                user:
                displayName: "Patti Fernandez"
                email: "PattiF@5bsjy7.onmicrosoft.com"
                id: "259baaaa-f37c-4e51-993f-782d71fe6005"
                [[Prototype]]: Object
                [[Prototype]]: Object
                lastModifiedDateTime: "2022-03-10T09:43:52Z"
                name: "New Microsoft Word Document (2).docx"
                parentReference: {driveId: 'b!uISMFsVEOk6w4T_lIDdFQ5caxuKNRo1GgEzwmh8I8Jz3RHo8T8Z-Q7f6WHSGW4yt', driveType: 'documentLibrary', id: '01JZT46KN6Y2GOVW7725BZO354PWSELRRZ', path: '/drive/root:'}
                size: 17865
                webUrl: "https://lingsuns.sharepoint.com/sites/SampleTeamSite/_layouts/15/Doc.aspx?sourcedoc=%7B94251CD1-4A09-4F49-A4AD-DDDA14654439%7D&file=New%20Microsoft%20Word%20Document%20(2).docx&action=default&mobileredirect=true"
                */
            
                var nextLink = "";
                var deltaLink = ""; 

                do
                {
                    try
                    {
                        var files = await GraphAPIHelper.CheckForUpdates(this.props.msGraphClient,nextLink,siteID,this.state.queryStartDate,drives[k]['id']);
                    }
                    catch
                    {
                        
                        this.setState({queried:true,
                        message:`${strings.FC_Ex_GetFilesChangeError}`,         
                        messageType:MessageBarType.error
                        });
                        SPOQASpinner.Hide();
                        return;
                    }
                    console.log(files);
                    for(var i=1; i<files.value.length; i++)
                    {
                        try{
                            if(typeof files.value[i]['deleted'] !== 'undefined')
                            {
                                console.log(files.value[i]);
                            }
                            else
                            {
                                let aFile:IFile = {
                                    ModifiedByEmail: "",
                                    ModifiedByName:"",
                                    ModifiedDate:"",
                                    Path:"",
                                    //Id:"",
                                    FileName:"",
                                    Library:""
                                };
                                
                                /*
                                ModifiedByEmail:string;
                                ModifiedByName:string;
                                ModifiedDate:string;
                                Path:string;
                                FileName:string;
                                Id:string;
                                */
                                console.log(files.value[i]);
                                aFile['ModifiedByEmail'] = `${files.value[i]['lastModifiedBy']['user']['email']}`;
                                aFile['ModifiedByName'] = `${files.value[i]['lastModifiedBy']['user']['displayName']}`;
                                aFile['ModifiedDate'] = `${files.value[i]['lastModifiedDateTime']}`;
                                aFile['Path'] = `${files.value[i]['webUrl']}`;
                                aFile['FileName'] = `${files.value[i]['name']}`;
                                //aFile['Id'] = `${files.value[i]['id']}`;
                                aFile['Library'] = `${drives[k]['name']}`;

                                if(this.IsMatchFilter(aFile))
                                {
                                    this.modifiedFiles.push(aFile);
                                }
                            }
                        }
                        catch(error){
                            SPOQAHelper.ShowMessageBar("Error", `${error}`);
                        }
                    }
                    if(files['@odata.nextLink'])
                    {
                        nextLink = files['@odata.nextLink'];
                    }
                    else
                    {
                        deltaLink = files['@odata.deltaLink'];
                    }
                }while (deltaLink.length == 0);
            }
            
            // Get Deleted Files
            /*
            @odata.editLink: "SP.ChangeItem33a91460-981a-42b0-8a1d-861fd05778cf"
            @odata.id: "https://lingsuns.sharepoint.com/sites/29738881/_api/SP.ChangeItem33a91460-981a-42b0-8a1d-861fd05778cf"
            @odata.type: "#SP.ChangeItem"
            ChangeToken:
            StringValue: "1;1;01ed74ae-3f05-41fd-a81a-47359ecb3178;637840550097370000;36674898"
            [[Prototype]]: Object
            ChangeType: 3
            Editor: ""
            EditorEmailHint: null
            ItemId: 1
            ListId: "4f8fa20d-d415-4de9-95d3-5b32451ed8b8"
            ServerRelativeUrl: ""
            SharedByUser: null
            SharedWithUsers: null
            SiteId: "01ed74ae-3f05-41fd-a81a-47359ecb3178"
            Time: "2022-03-28T09:03:30Z"
            UniqueId: "b7a1ad48-e436-4970-ae51-8b4ad821b74d"
            WebId: "359810b0-b65b-480b-b33f-a9c4dd200f4b"
            */

            /*
            var files2 = await RestAPIHelper.GetSiteChanges(this.props.spHttpClient,siteID,this.state.querySite,this.state.queryStartDate);
            console.log(files2);

            for(var j=0; j<files2.value.length; j++)
            {
                let bFile:IFile = {
                    ModifiedByEmail: "",
                    ModifiedByName:"",
                    ModifiedDate:"",
                    Path:"",
                    //Id:"",
                    FileName:"",
                    Library:""
                };

                var listUrl = await RestAPIHelper.GetListPath(this.props.spHttpClient,this.state.querySite,files2.value[j]['ListId']);
                var list = await RestAPIHelper.GetListbyId(this.props.spHttpClient,this.state.querySite,files2.value[j]['ListId']);

                bFile['ModifiedByEmail'] = ``;
                bFile['ModifiedByName'] = ``;
                bFile['ModifiedDate'] = `${files2.value[j]['Time']}`;
                bFile['Path'] = `${this.props.rootUrl}${listUrl}`;
                bFile['FileName'] = `<deleted file>`;
                //bFile['Id'] = `${files2.value[j]['UniqueId']}`;
                bFile['Library'] = `${list['Title']}`;

                if(this.IsMatchFilter2(bFile))
                {
                    this.modifiedFiles.push(bFile);
                }
            }…*/

            this.setState({queried:true,
            message:`${strings.FC_Message_QueryDone}  ${this.modifiedFiles.length}`,
            messageType:MessageBarType.success
            });
            
            SPOQASpinner.Hide();
        }
        else
        {
            SPOQAHelper.ShowMessageBar("Error", `${strings.FC_Ex_GetSiteError} ${this.state.querySite}!`);
        }
    }

    private IsMatchFilter(item:any):boolean
    {
        let matched:boolean = true;
        if(this.state.modifiedByUser && this.state.modifiedByUser.trim().length >0)
        {
            matched = matched&&(this.state.modifiedByUser.toLowerCase() == item.ModifiedByEmail.toLowerCase());
        }
        if(this.state.pathFilter && this.state.pathFilter.trim().length > 0)
        {
            matched = matched&&((item.Path.toLowerCase().indexOf(this.state.pathFilter.toLowerCase()) >= 0)||(item.Library.toLowerCase().indexOf(this.state.pathFilter.toLowerCase()) >= 0));
        }
        if(this.state.queryEndDate)
        {
            let queryEndDate:Date = new Date(this.state.queryEndDate);
            queryEndDate.setDate(queryEndDate.getDate()+1);
            matched = matched&&(queryEndDate >= new Date(item.ModifiedDate));
        }
        return matched;
    }

    /*
    private IsMatchFilter2(item:any):boolean
    {
        let matched:boolean = true;

        if(this.state.pathFilter && this.state.pathFilter.trim().length > 0)
        {
            matched = matched&&((item.Path.toLowerCase().indexOf(this.state.pathFilter.toLowerCase()) >= 0||(item.Library.toLowerCase().indexOf(this.state.pathFilter.toLowerCase())) >= 0));
        }
        if(this.state.queryEndDate)
        {
            let queryEndDate:Date = new Date(this.state.queryEndDate);
            queryEndDate.setDate(queryEndDate.getDate()+1);
            matched = matched&&(queryEndDate >= new Date(item.ModifiedDate));
        }
        return matched;
    }*/
    
    private DoExport():void
    {
        var ts = new Date().getTime();
        // Export filtered files
        SPOQAHelper.JSONToCSVConvertor(this.modifiedFiles, true, `FileDelta_${ts}`);
    }
}

