import * as React from 'react'; 
import styles from '../SharePointOnlineQuickAssist.module.scss';
import {  
    PrimaryButton,
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
                                    label="Query Site:"
                                    multiline={false}
                                    onChange={(e)=>{let text:any = e.target; this.setState({querySite:text.value});}}
                                    value={this.state.querySite}
                                    required={true}                        
                                /> 
                            </div>
                            <div className={styles.msrow} id="queryFilter_row">
                                <div className={styles.mscol6}>
                                    <TextField
                                        label="Modified User:"                        
                                        multiline={false}
                                        onChange={(e)=>{let text:any = e.target; this.setState({modifiedByUser:text.value});}}
                                        value={this.state.modifiedByUser}
                                    />
                                    </div>          
                                    <div className={styles.mscol6}>
                                    <TextField 
                                        label="Path Filter:"
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
                                        label='Start Date:'
                                        placeholder="Select a date..."
                                        ariaLabel="Select a date"
                                        onSelectDate={(e)=>{ this.setState({queryStartDate:e});}}
                                        value={this.state.queryStartDate}                    
                                    />
                                </div>
                                <div className={styles.mscol6}>
                                    <DatePicker
                                        label='End Date:'
                                        placeholder="Select a date..."
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
                            <PrimaryButton
                                text="Get Files"
                                style={{ display: 'inline', marginTop: '10px' }}
                                onClick={() => {SPOQAHelper.ResetFormStaus();this.QueryFiles();}}
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
        //Get Site ID
        var siteID = await RestAPIHelper.GetSiteId(this.props.spHttpClient, this.state.querySite);
        if(siteID)
        {   
            this.modifiedFiles = [];
            this.setState({queried:false});
            SPOQASpinner.Show("Querying ......");
            
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
                    var files = await GraphAPIHelper.CheckForUpdates(this.props.msGraphClient,nextLink,siteID,this.state.queryStartDate);
                }
                catch
                {
                    
                    this.setState({queried:true,
                    message:`Get Files Change Exited Unexpectedly`,         
                    messageType:MessageBarType.error
                    });
                    SPOQASpinner.Hide();
                    return;
                }
                console.log(files);
                for(var i=1; i<files.value.length; i++)
                {
                    try{
                        let aFile:IFile = {
                            ModifiedByEmail: "",
                            ModifiedByName:"",
                            ModifiedDate:"",
                            Path:"",
                            Id:"",
                            FileName:""
                        };

                        if(typeof files.value[i]['deleted'] !== 'undefined')
                        {
                            console.log(files.value[i]);
                            aFile['FileName'] = `deleted file`;
                        }
                        else
                        {
                            if(this.IsMatchFilter(files.value[i]))
                            {
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
                            }
                        }
                        aFile['Id'] = `${files.value[i]['id']}`;
                        this.modifiedFiles.push(aFile);
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

            this.setState({queried:true,
            message:`Query Complete. Changes Number: ${this.modifiedFiles.length}`,
            messageType:MessageBarType.success
            });
            SPOQASpinner.Hide();
        }
        else
        {
            SPOQAHelper.ShowMessageBar("Error", `Failed to get the site ${this.state.querySite}!`);
        }
    }

    private IsMatchFilter(item:any):boolean
    {
        let matched:boolean = true;
        if(this.state.modifiedByUser && this.state.modifiedByUser.trim().length >0)
        {
        matched = matched&&(this.state.modifiedByUser.toLowerCase() == item.lastModifiedBy.user.email.toLowerCase());
        }
        if(this.state.pathFilter && this.state.pathFilter.trim().length > 0)
        {
        matched = matched&&(item.webUrl.toLowerCase().indexOf(this.state.pathFilter.toLowerCase()) >= 0);
        }
        if(this.state.queryEndDate)
        {
            let queryEndDate:Date = new Date(this.state.queryEndDate);
            queryEndDate.setDate(queryEndDate.getDate()+1);
            matched = matched&&(queryEndDate >= new Date(item.lastModifiedDateTime));
        }
        return matched;
    }

}

