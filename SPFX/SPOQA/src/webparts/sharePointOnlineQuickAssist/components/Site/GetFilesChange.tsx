import * as React from 'react'; 
import styles from '../SharePointOnlineQuickAssist.module.scss';
import {  
    PrimaryButton,
    MessageBar,
    MessageBarType
  } from 'office-ui-fabric-react/lib/index';
  import { ISharePointOnlineQuickAssistProps } from '../ISharePointOnlineQuickAssistProps';
import GraphAPIHelper from '../../../Helpers/GraphAPIHelper';  
import SPOQASpinner from '../../../Helpers/SPOQASpinner';
import SPOQAHelper from '../../../Helpers/SPOQAHelper';
import IFilesGrid from "./IFilesGrid";
import { IFile } from "./IFile";
export default class GetFilesChange extends React.Component<ISharePointOnlineQuickAssistProps>
{
    private modifiedFiles:IFile[];    
    public state = {
        message:"",
        messageType:MessageBarType.success,
        queried:false
    };

    public render():React.ReactElement<ISharePointOnlineQuickAssistProps>
    {
        return (
          <div>
            <div id="IFiles_CommandButtonsSection">
                  <PrimaryButton
                      text="Get FIles"
                      style={{ display: 'inline', marginTop: '10px' }}
                      onClick={() => {this.QueryFiles();}}
                    />
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
              var files = await GraphAPIHelper.CheckForUpdates(this.props.msGraphClient,nextLink);
              console.log(files);
              for(var i=1; i<files.value.length; i++)
              {
                let aFile:IFile = {
                  ModifiedByEmail: "",
                  ModifiedByName:"",
                  ModifiedDate:"",
                  Path:"",
                  Id:"",
                  FileName:""
                };
                /*
                    ModifiedByEmail:string;
                    ModifiedByName:string;
                    ModifiedDate:string;
                    Path:string;
                    FileName:string;
                    Id:string;
                */
                aFile['ModifiedByEmail'] = `${files.value[i]['lastModifiedBy']['user']['email']}`;
                aFile['ModifiedByName'] = `${files.value[i]['lastModifiedBy']['user']['displayName']}`;
                aFile['ModifiedDate'] = `${files.value[i]['lastModifiedDateTime']}`;
                aFile['Path'] = `${files.value[i]['webUrl']}`;
                aFile['FileName'] = `${files.value[i]['name']}`;
                aFile['Id'] = `${files.value[i]['id']}`;
                this.modifiedFiles.push(aFile);

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

}

