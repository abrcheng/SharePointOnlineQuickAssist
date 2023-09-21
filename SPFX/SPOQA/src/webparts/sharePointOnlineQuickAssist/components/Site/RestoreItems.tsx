import * as React from 'react';
import { DateTimePicker,DateConvention, TimeConvention, TimeDisplayControlType} from '@pnp/spfx-controls-react';

import {     
    DefaultButton,     
    TextField,
    MessageBar,
    MessageBarType,
    //DatePicker,
    Spinner,
    Toggle    
  } from 'office-ui-fabric-react/lib/index';
import RestAPIHelper from '../../../Helpers/RestAPIHelper';
import SPOQASpinner from '../../../Helpers/SPOQASpinner';
import SPOQAHelper from '../../../Helpers/SPOQAHelper';
import RestoreItemsQAGrid from "./RestoreItemsQAGrid";
import styles from '../SharePointOnlineQuickAssist.module.scss';
import { ISharePointOnlineQuickAssistProps } from '../ISharePointOnlineQuickAssistProps';
import { IRestoreItem,IRestoreItems } from "./IRestoreItem";
import { Text } from '@microsoft/sp-core-library';
import * as strings from 'SharePointOnlineQuickAssistWebPartStrings';

export default class RestoreItemsQA extends React.Component<ISharePointOnlineQuickAssistProps>
{
  private recycleBinItems:IRestoreItem[];
  private queryCount:number = 0;
  private querySeconds:number=0;

  public state = {
    deleteByUser: this.props.currentUser.loginName,
    deleteStartDate:null,
    deleteEndDate:null,
    pathFilter:"",    
    affectedSite:this.props.webAbsoluteUrl,
    queried:false,
    currentItems:[],
    message:"",
    messageType:MessageBarType.success,
    spinnerMessage:"",
    errorDetail:[],
    detectAndSkipExistingDocument:false
  };
  
  // https://chengc.sharepoint.com/sites/abc/_api/site/getrecyclebinitems?rowLimit='100'&isAscending=false&itemState=1&orderby=3
  // 'id=dbe08209-a916-4762-8390-200aeefe91f2&title=Table of Contents.docx&searchValue=2021-12-21T08:25:47' => encode => pagingInfo
  // https://chengc.sharepoint.com/sites/abc/_api/site/getrecyclebinitems?rowLimit=%27101%27&isAscending=false&itemState=1&orderby=3&pagingInfo=%27id%3Ddbe08209-a916-4762-8390-200aeefe91f2%26title%3DTable%20of%20Contents.docx%26searchValue%3D2021-12-21T08%3A25%3A47%27
  // <d:DeletedByEmail>abc@chengc.onmicrosoft.com</d:DeletedByEmail>
  // <d:DeletedByName>Abraham  Cheng</d:DeletedByName>
  // <d:DeletedDate m:type="Edm.DateTime">2022-02-09T02:31:10Z</d:DeletedDate>
  // <d:LeafName>desktop.ini</d:LeafName>
  // <d:DirName>sites/abc/Shared Documents/Program Files</d:DirName>
  // <d:Id m:type="Edm.Guid">fa68c9fc-0c8e-4d50-ade7-def556523bb1</d:Id>
  public render():React.ReactElement<ISharePointOnlineQuickAssistProps>
  {
      return (
        <div>
          <div className={ styles.row }>
            <div className={ styles.column }>
            <div id="RestoreItemsQA_FilterSection" className={styles.msgrid}>
              <div className={styles.msrow} id="affectedSite_row">
                <div className={styles.mscol8}>
                  <TextField
                        label={strings.AffectedSite}
                        multiline={false}
                        onChange={(e)=>{let text:any = e.target; this.setState({affectedSite:text.value});}}
                        value={this.state.affectedSite}
                        required={true}                        
                  /> 
                </div>
                <div className={styles.mscol4}>
                  <Toggle 
                  label={strings.RI_DetectAndSkipExistingDocument}
                  onChange={(e, checked)=>{ this.setState({detectAndSkipExistingDocument:checked});}}
                  onText="On"
                  offText="Off"
                  checked={this.state.detectAndSkipExistingDocument}/>                                        
                </div>
              </div>
              <div className={styles.msrow} id="deleteByUser_row">
              <div className={styles.mscol6}>
                <TextField
                            label={strings.RI_DeletedBy}                       
                            multiline={false}
                            onChange={(e)=>{let text:any = e.target; this.setState({deleteByUser:text.value});}}
                            value={this.state.deleteByUser}
                            
                      />
                </div>          
                <div className={styles.mscol6}>
                  <TextField 
                              label={strings.RI_PathFilter}
                              className='ms-Grid-col ms-u-sm6 block'
                              multiline={false}
                              onChange={(e)=>{let text:any = e.target; this.setState({pathFilter:text.value});}}
                              value={this.state.pathFilter}       
                                                 
                    />    
                </div>    
              </div>
              
              <div className={styles.msrow} id="deleteStartDate_row">
                    <div className={styles.mscol6}>
                      
                      <DateTimePicker 
                         dateConvention={DateConvention.DateTime}
                         timeConvention={TimeConvention.Hours24} 
                         showSeconds= {true}
                         timeDisplayControlType={TimeDisplayControlType.Dropdown}                           
                         // isMonthPickerVisible ={false}                              
                          label={strings.RI_StartDate}
                          placeholder={strings.RI_SelectADate}                          
                          // onChange={(e)=>{let datePicker:any = e.target; this.setState({deleteStartDate:datePicker.value});}}
                          onChange={(e)=>{ this.setState({deleteStartDate:e});}}
                          value={this.state.deleteStartDate}                    
                      />
                    </div>
                    <div className={styles.mscol6}>
                      <DateTimePicker
                         dateConvention={DateConvention.DateTime}
                         timeConvention={TimeConvention.Hours24}
                         timeDisplayControlType={TimeDisplayControlType.Dropdown}   
                         showSeconds= {true}                                             
                          // isMonthPickerVisible={false}     
                          label={strings.RI_EndDate}
                          placeholder={strings.RI_SelectADate}                          
                          // onChange={(e)=>{let datePicker:any = e.target; this.setState({deleteEndDate:datePicker.value});}}
                          onChange={(e)=>{ this.setState({deleteEndDate:e});}}
                          value={this.state.deleteEndDate}                                                    
                      />
                    </div>
                </div>             
              </div>
            </div>
            </div>

            <div className={ styles.row }>
            <div className={ styles.column }>
              <div id="RestoreItemsQA_CommandButtonsSection">
                  <DefaultButton
                            text={strings.RI_QueryItems}
                            style={{ display: 'inline', marginTop: '10px' }}
                            
                            onClick={() => {this.QueryRecycleBinItems();}}
                            />
                  {this.state.queried && this.state.currentItems.length >0?
                            <div style={{ display: 'inline'}}>
                            <DefaultButton
                                text={strings.RI_Restore}
                                style={{ display: 'inline', marginTop: '10px', marginLeft:"10px"}}
                                onClick={() => {this.Restore();}}
                            /> 
                               <DefaultButton
                                text= {strings.RI_Export}
                                style={{ display: 'inline', marginTop: '10px', marginLeft:"10px"}}
                                onClick={() => {this.DoExport();}}
                            />
                           </div>:null}                  
                  </div>
                </div>
              </div>

          <div id="RestoreItemsQA_QueryResultSection">
              {this.state.spinnerMessage !=""? <Spinner id="SPOQASpinner" label={this.state.spinnerMessage} ariaLive="assertive" labelPosition="left" />:null}
              {this.state.queried?<MessageBar id="RestoreItemsQAMessageBar" messageBarType={this.state.messageType} isMultiline={true}>
                 {this.state.message}
              </MessageBar>:null}
              {this.state.errorDetail.length >0?<div>{this.state.errorDetail.map(error => <div style={{color:"red"}}>{error}</div>)}</div>:null}
              {this.state.queried && this.state.currentItems.length >0? <RestoreItemsQAGrid items={this.state.currentItems}/>:null}
          </div>          
          <div id="RestoreItemsQA_ActionSection">
          </div>
        </div>
      );
  }

  private async QueryRecycleBinItems()
  {
     // Verify the site box is null 
    if(this.state.affectedSite =="")
    {
      SPOQAHelper.ShowMessageBar("Error", strings.UI_NonAffectedSite);          
      return;
    }
  else{
    
      // Verify the site is valid 
      this.setState({errorDetail:[],currentItems:[]});
      var isSiteValid = await RestAPIHelper.GetWeb(this.props.spHttpClient, this.state.affectedSite);
      if(isSiteValid)
      {   
          var pageInfo = "";
          var currentCount=-1;
          this.recycleBinItems = []; // clean previous data set
          this.setState({queried:false});
          SPOQASpinner.Show(`${strings.RI_Querying} ......`);
          var itemState = 1;
          this.queryCount = 0;
          var queryStartTime = new Date();

          while(currentCount ==500 || currentCount==-1 || itemState <3)
          {
            var recycleItems = await RestAPIHelper.Getrecyclebinitems(this.props.spHttpClient, this.state.affectedSite,pageInfo,500,false, itemState, 3);
            if(recycleItems.error)
            {
              SPOQAHelper.ShowMessageBar("Error",Text.format(strings.RI_GetrecyclebinitemsError, this.state.affectedSite, recycleItems.error.message));
              SPOQASpinner.Hide();
              return;
            }

            // recycleItems.value.length, if the length is less than 500, that's mean the current query is the last page
            /* Data structure of recycleItems.value[0]
              @odata.editLink: "Site/RecycleBin(guid'ca1352a7-2124-4a93-8254-748554761319')"
              @odata.id: "https://chengc.sharepoint.com/sites/abc/_api/Site/RecycleBin(guid'ca1352a7-2124-4a93-8254-748554761319')"
              @odata.type: "#SP.RecycleBinItem"
              AuthorEmail: ""
              AuthorName: "System Account"
              DeletedByEmail: "abc@chengc.onmicrosoft.com"
              DeletedByName: "Abraham  Cheng"
              DeletedDate: "2022-02-19T10:16:51Z"
              DeletedDateLocalFormatted: "2/19/2022 2:16 AM"
              DirName: "sites/abc/ABPMicroService/Acme.BookStore/react-native/node_modules/compression"
              DirNamePath: {DecodedUrl: 'sites/abc/ABPMicroService/Acme.BookStore/react-native/node_modules/compression'}
              Id: "ca1352a7-2124-4a93-8254-748554761319"
              ItemState: 1
              ItemType: 1
              LeafName: "LICENSE"
              LeafNamePath: {DecodedUrl: 'LICENSE'}
              Size: 1563
              Title: "LICENSE"
            */   
            var lastId = "";
            var lastTitle ="";
            var lastDeletedDate = "";
            currentCount = recycleItems.value.length;
            for(var i=0; i<recycleItems.value.length; i++)
            {
                var currentItem = recycleItems.value[i];
                if(lastId != currentItem.Id)
                {
                  this.queryCount ++;
                  currentItem["Path"] = `${currentItem["DirName"]}/${currentItem["LeafName"]}`;
                  delete currentItem['DirName'];
                  delete currentItem['LeafName'];
                  if(this.IsMatchFilter(currentItem))
                  {
                    delete currentItem['LeafNamePath'];
                    delete currentItem['DirNamePath'];
                    delete currentItem['@odata.editLink'];
                    delete currentItem['@odata.id'];
                    delete currentItem['@odata.type'];
                    // If the "Detect and skip existing items" is off, then Existing will always set to false by default
                    //  If the "Detect and skip existing items" is on, the Existing will be filled latter in the function DetectExistingItems
                    currentItem["Existing"] = false;     
                    this.recycleBinItems.push(currentItem);
                  }
                }

                lastId = currentItem.Id;
                lastTitle = currentItem.Title;
                lastDeletedDate = currentItem.DeletedDate;
            }
              
            console.log(`RestAPIHelper.Getrecyclebinitems with page info ${pageInfo}`);
            if(lastDeletedDate.indexOf("Z") > 0)
            {
                lastDeletedDate = lastDeletedDate.substring(0, lastDeletedDate.length -1);
            }

            pageInfo = URI_Encoding.encodeURIComponent(`'id=${URI_Encoding.encodeURIComponent(lastId)}&title=${URI_Encoding.encodeURIComponent(lastTitle)}&searchValue=${URI_Encoding.encodeURIComponent(lastDeletedDate)}'`); 
            
            SPOQASpinner.Show(Text.format(strings.RI_QueryProgress, this.queryCount, this.recycleBinItems.length));  
            if(currentCount <500)    
            {
                if(itemState ==1)
                {
                  itemState = 2;
                  pageInfo = "";
                }
                else
                {
                  itemState =3;
                }
            } 
                     
         }
        
         this.querySeconds = ((new Date()).getTime()- queryStartTime.getTime())/1000;
         this.recycleBinItems.sort((a,b) =>a.Path > b.Path ?1:-1);   
         this.setState({
              message: Text.format(strings.RI_QueryResult, this.queryCount, this.recycleBinItems.length, this.querySeconds),              
              messageType:MessageBarType.success,
              queried:true
          });  
          
         if(this.state.detectAndSkipExistingDocument) // Detect existing items
         {
            var dtected = await this.DetectExistingItems(this.state.affectedSite);
            if(dtected)
            {
              var existingItemCount = 0;
            this.recycleBinItems.forEach(item=>{
                if(item.Existing)
                {
                  existingItemCount++;
                }
            });

            this.setState({currentItems:this.recycleBinItems,
                  message: `${Text.format(strings.RI_QueryResult, this.queryCount, this.recycleBinItems.length, this.querySeconds)} ${Text.format(strings.RI_DetectExistingResult, existingItemCount)}`,
                  queried:true,
                  messageType:MessageBarType.success
              }); 
            }
         }
         else
         {
          this.setState({currentItems:this.recycleBinItems});
         }

         SPOQASpinner.Hide();
     }
     else
     {
        SPOQAHelper.ShowMessageBar("Error", `Failed to get the site ${this.state.affectedSite}!`);
     }
    }
  }

  private IsMatchFilter(item:any):boolean
  {
    let matched:boolean = true;
    if(this.state.deleteByUser && this.state.deleteByUser.trim().length >0) // check deleteByUser
    {
      if(!item.DeletedByEmail) // If the DeletedByEmail is null, then set it to empty string 
      {
        item.DeletedByEmail="";
      }

      if(!item.DeletedByName) // If the DeletedByName is null, then set it to empty string 
      {
        item.DeletedByName="";
      }

      matched = matched&&((item.DeletedByEmail.toLowerCase().indexOf(this.state.deleteByUser.trim().toLowerCase()) >= 0)||(item.DeletedByName.toLowerCase().indexOf(this.state.deleteByUser.trim().toLowerCase()) >= 0));
      //matched = matched&&(this.state.deleteByUser.trim().toLowerCase() == item.DeletedByEmail.toLowerCase());
    }

    if(this.state.pathFilter && this.state.pathFilter.trim().length > 0)
    {
      matched = matched&&(item.Path.toLowerCase().indexOf(this.state.pathFilter.trim().toLowerCase()) >= 0);
    }

    if(this.state.deleteStartDate)
    {
       let deleteStartDate:Date = new Date(this.state.deleteStartDate);
       deleteStartDate.setMinutes(deleteStartDate.getMinutes() - deleteStartDate.getTimezoneOffset());
       matched = matched&&(deleteStartDate <= new Date(item.DeletedDate));
    }
    
    if(this.state.deleteEndDate)
    {
        let deleteEndDate:Date = new Date(this.state.deleteEndDate);
        deleteEndDate.setMinutes(deleteEndDate.getMinutes() - deleteEndDate.getTimezoneOffset());
        // deleteEndDate.setDate(deleteEndDate.getDate()+1);
        matched = matched&&(deleteEndDate >= new Date(item.DeletedDate));
    }
   
    /*
    const date = new Date();    
    date.setMinutes(date.getMinutes() - date.getTimezoneOffset())
    1695308870444
    date.toGMTString()
    'Thu, 21 Sep 2023 15:07:50 GMT'
    */
    return matched;
  }

  private async Restore()
  {    
    // https://chengc.sharepoint.com/sites/SPOQA/_api/site/RecycleBin/RestoreByIds post
    // {"ids":
    //    ["59def901-c4b2-433a-ac74-5c2fbc5dd933",
    //     "f2621b7b-7732-4423-95d0-60321e80fa65"]
    // ,"bRenameExistingItems":true}
    
    // Restore 10 items in one batch
    var  recycleBinItemsNeedRestore = this.recycleBinItems.filter(e=>!e.Existing);
    var batchSize = 10;
    let batchNo:number = Math.ceil(recycleBinItemsNeedRestore.length /batchSize);
    var restoreStartTime = new Date();
    this.setState({errorDetail:[]});
    for(var batchIndex=0; batchIndex <batchNo;batchIndex++)
    {
      let ids:string[]=[];
      let startIndex:number= batchIndex * batchSize;
      let endIndex:number = (batchIndex+1) * batchSize < recycleBinItemsNeedRestore.length? (batchIndex+1) * batchSize : recycleBinItemsNeedRestore.length;
      for(var index=startIndex; index < endIndex; index++)
      {
        ids.push(recycleBinItemsNeedRestore[index].Id);
      }
     
      this.setState({     
        spinnerMessage: Text.format(strings.RI_RestoreProgress, startIndex + 1, endIndex)
        });  

      let restoreRes = await RestAPIHelper.RestoreByIds(this.props.spHttpClient, this.state.affectedSite, ids);
      if(restoreRes.success)
      {         
          if(batchIndex + 1 == batchNo) // last bacth completed
          {
            var restoreSeconds = ((new Date()).getTime()- restoreStartTime.getTime())/1000;
            this.setState({
              message: Text.format(strings.RI_RestoreResult, recycleBinItemsNeedRestore.length, restoreSeconds, this.recycleBinItems.length -  recycleBinItemsNeedRestore.length),         
              messageType:MessageBarType.success,
              spinnerMessage:""
              }); 
          }
          else
          {
            /*this.setState({
              message:`${strings.RI_RestoreItemFrom} ${startIndex + 1} ${strings.RI_To} ${endIndex}.`,         
              messageType:MessageBarType.success,
              spinnerMessage:""
              });*/  
          }
      }
      else
      {
        // restoreRes.error.message        
        const { errorDetail} = this.state;  
        errorDetail.push(restoreRes.error.message);
        this.setState({
          message:Text.format(strings.RI_RestoreResultWithError,startIndex, endIndex, restoreRes.error.message),         
          messageType:MessageBarType.error,
          spinnerMessage:"",
          errorDetail:errorDetail
          });          
      }
    }
  }

  private DoExport():void
  {
      // Export filtered recycle bin items
      SPOQAHelper.JSONToCSVConvertor(this.recycleBinItems, true, "RecycleBinItems");
  }

  private async DetectExistingItems(webUrl:string) {
    // https://chengc.sharepoint.com/sites/SPOQA/_api/web/webs?$select=Url
    SPOQASpinner.Show(Text.format(strings.RI_DetectExistingItemsInPath, webUrl));
    var lists:any = await RestAPIHelper.GetLists(this.props.spHttpClient, webUrl); 
    for(var listIndex=0; listIndex < lists.length;listIndex++)
    {
      var list = lists[listIndex];
         // rootFolder:list.RootFolder.ServerRelativeUrl 
      if(list.RootFolder.ServerRelativeUrl.indexOf("/") == 0)
      {
        list.RootFolder.ServerRelativeUrl = list.RootFolder.ServerRelativeUrl.substring(1,list.RootFolder.ServerRelativeUrl.length);
      }
      
      var hasItemMatchList = this.recycleBinItems.some(e=>e.Path.toLowerCase().indexOf(list.RootFolder.ServerRelativeUrl.toLowerCase()) ===0);                         
      if(hasItemMatchList)
      {
        // Matched some items in current list by list.RootFolder.ServerRelativeUrl , detect existing items in the current list
        await this.DetectExistingItemsInList(webUrl,list);
      }
    }
   
    // Get sub sites and call DetectExistingItems for each sub site
    var subWebs = await RestAPIHelper.GetSubWebs(this.props.spHttpClient, webUrl);
  
    if(subWebs && subWebs.value && subWebs.value.length >0)
    {
      for(var webIndex=0; webIndex < subWebs.value.length; webIndex++)
      {
        var url = subWebs.value[webIndex].ServerRelativeUrl.substring(1,subWebs.value[webIndex].ServerRelativeUrl.length);
        var hasItemMatchWeb = this.recycleBinItems.some(e=>e.Path.toLowerCase().indexOf(url.toLowerCase()) ===0);   
        if(hasItemMatchWeb)
        {
          await this.DetectExistingItems(subWebs.value[webIndex].Url);
        }
      }     
    }

    return await subWebs;
  }
  
  private async DetectExistingItemsInList(webUrl:string, list:any) { 
    SPOQASpinner.Show(Text.format(strings.RI_DetectExistingItemsInPath, list.RootFolder.ServerRelativeUrl));
    let allItemsList:any[] = await RestAPIHelper.GetItemsInList(this.props.spHttpClient, webUrl, list.Id);
    console.log(`Get ${allItemsList.length} items from list ${list.RootFolder.ServerRelativeUrl}`);
    allItemsList.push({FileRef:list.RootFolder.ServerRelativeUrl});
    var matchCount = 0;
    allItemsList.forEach(item=>{
      if(item.FileRef.indexOf("/")===0)
      {
        item.FileRef = item.FileRef.substring(1, item.FileRef.length);
      }

      for(var index=0; index<this.recycleBinItems.length; index++)
      {
        if(item.FileRef.toLowerCase()=== this.recycleBinItems[index].Path.toLowerCase())
        {
          this.recycleBinItems[index].Existing = true;
          matchCount++;
        }
      }
    });
    
    console.log(`Get ${matchCount} items matched in the list ${list.RootFolder.ServerRelativeUrl}`);
  }
}