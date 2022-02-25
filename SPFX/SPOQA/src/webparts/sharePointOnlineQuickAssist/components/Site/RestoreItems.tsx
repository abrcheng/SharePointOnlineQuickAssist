import * as React from 'react';
import {  
    PrimaryButton,    
    TextField,
    MessageBar,
    MessageBarType,
    DatePicker
  } from 'office-ui-fabric-react/lib/index';
import GraphAPIHelper from '../../../Helpers/GraphAPIHelper';  
import RestAPIHelper from '../../../Helpers/RestAPIHelper';
import SPOQASpinner from '../../../Helpers/SPOQASpinner';
import SPOQAHelper from '../../../Helpers/SPOQAHelper';
import RestoreItemsQAGrid from "./RestoreItemsQAGrid";
import styles from '../SharePointOnlineQuickAssist.module.scss';
import { ISharePointOnlineQuickAssistProps } from '../ISharePointOnlineQuickAssistProps';
import { IRestoreItem,IRestoreItems } from "./IRestoreItem";
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
    currentItems:null,
    message:"",
    messageType:MessageBarType.success
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
                <TextField
                      label="Affected Site:"
                      multiline={false}
                      onChange={(e)=>{let text:any = e.target; this.setState({affectedSite:text.value});}}
                      value={this.state.affectedSite}
                      required={true}                        
                /> 
              </div>
              <div className={styles.msrow} id="deleteByUser_row">
              <div className={styles.mscol6}>
                <TextField
                            label="Deleted User:"                        
                            multiline={false}
                            onChange={(e)=>{let text:any = e.target; this.setState({deleteByUser:text.value});}}
                            value={this.state.deleteByUser}
                            
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
              <div className={styles.msrow} id="deleteStartDate_row">
                    <div className={styles.mscol6}>
                      <DatePicker
                          label='Start Date:'
                          placeholder="Select a date..."
                          ariaLabel="Select a date"   
                          // onChange={(e)=>{let datePicker:any = e.target; this.setState({deleteStartDate:datePicker.value});}}
                          onSelectDate={(e)=>{ this.setState({deleteStartDate:e});}}
                          value={this.state.deleteStartDate}                    
                      />
                    </div>
                    <div className={styles.mscol6}>
                      <DatePicker
                          label='End Date:'
                          placeholder="Select a date..."
                          ariaLabel="Select a date"   
                          // onChange={(e)=>{let datePicker:any = e.target; this.setState({deleteEndDate:datePicker.value});}}
                          onSelectDate={(e)=>{ this.setState({deleteEndDate:e});}}
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
                  <PrimaryButton
                            text="Query Items"
                            style={{ display: 'inline', marginTop: '10px' }}
                            onClick={() => {this.QueryRecycleBinItems();}}
                            />
                  {this.state.queried && this.state.currentItems.length >0?
                            <div style={{ display: 'inline'}}>
                            <PrimaryButton
                                text="Restore"
                                style={{ display: 'inline', marginTop: '10px', marginLeft:"10px"}}
                                onClick={() => {this.Restore();}}
                            /> 
                               <PrimaryButton
                                text="Export"
                                style={{ display: 'inline', marginTop: '10px', marginLeft:"10px"}}
                                onClick={() => {this.DoExport();}}
                            />
                           </div>:null}                  
                  </div>
                </div>
              </div>

          <div id="RestoreItemsQA_QueryResultSection">
              {this.state.queried?<MessageBar id="RestoreItemsQAMessageBar" messageBarType={this.state.messageType} isMultiline={true}>
                 {this.state.message}
              </MessageBar>:null}
              {this.state.queried && this.state.currentItems.length >0? <RestoreItemsQAGrid items={this.state.currentItems}/>:null}
          </div>          
          <div id="RestoreItemsQA_ActionSection">
          </div>
        </div>
      );
  }

  private async QueryRecycleBinItems()
  {
     // Verify the site is valid 
     var isSiteValid = await RestAPIHelper.GetWeb(this.props.spHttpClient, this.state.affectedSite);
     if(isSiteValid)
     {   
         var pageInfo = "";
         var currentCount=-1;
         this.recycleBinItems = []; // clean previous data set
         this.setState({queried:false});
         SPOQASpinner.Show("Querying ......");
         var itemState = 1;
         this.queryCount = 0;
         var queryStartTime = new Date();

         while(currentCount ==500 || currentCount==-1 || itemState <3)
         {
           var recycleItems = await RestAPIHelper.Getrecyclebinitems(this.props.spHttpClient, this.state.affectedSite,pageInfo,500,false, itemState, 3);

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
                if(this.IsMatchFilter(currentItem))
                {
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
           SPOQASpinner.Show(`Queried ${this.queryCount} items, filtered ${this.recycleBinItems.length} items ......`);  
           if(currentCount <500)    
           {
              if(itemState ==1)
              {
                itemState = 2;
                pageInfo = "";
              }
              else
              {
                itemState =3
              }
           }
           
         }
        
         this.querySeconds = ((new Date()).getTime()- queryStartTime.getTime())/1000;
         this.recycleBinItems.sort((a,b) =>a.DirName > b.DirName ?1:-1);
         this.setState({currentItems:this.recycleBinItems,
              message:`Queried ${this.queryCount} items, filtered ${this.recycleBinItems.length} items in ${this.querySeconds} seconds.`,
              queried:true,
              messageType:MessageBarType.success
          });        
         SPOQASpinner.Hide();
     }
     else
     {
        SPOQAHelper.ShowMessageBar("Error", `Failed to get the site ${this.state.affectedSite}!`);
     }
  }

  private IsMatchFilter(item:any):boolean
  {
    let matched:boolean = true;
    if(this.state.deleteByUser && this.state.deleteByUser.trim().length >0) // check deleteByUser
    {
      matched = matched&&(this.state.deleteByUser.toLowerCase() == item.DeletedByEmail.toLowerCase());
    }

    if(this.state.pathFilter && this.state.pathFilter.trim().length > 0)
    {
      matched = matched&&(item.DirName.toLowerCase().indexOf(this.state.pathFilter.toLowerCase()) > 0);
    }

    if(this.state.deleteStartDate)
    {
        matched = matched&&(this.state.deleteStartDate <= new Date(item.DeletedDate));
    }
    
    if(this.state.deleteEndDate)
    {
        let deleteEndDate:Date = new Date(this.state.deleteEndDate);
        deleteEndDate.setDate(deleteEndDate.getDate()+1);
        matched = matched&&(deleteEndDate >= new Date(item.DeletedDate));
    }

    return matched;
  }

  private async Restore()
  {    
    // https://chengc.sharepoint.com/sites/SPOQA/_api/site/RecycleBin/RestoreByIds post
    // {"ids":
    //    ["59def901-c4b2-433a-ac74-5c2fbc5dd933",
    //     "f2621b7b-7732-4423-95d0-60321e80fa65"]
    // ,"bRenameExistingItems":true}
    
    // Restore 100 items in one batch 
    let batchNo:number = Math.ceil(this.recycleBinItems.length /100);
    for(var batchIndex=0; batchIndex <batchNo;batchIndex++)
    {
      let ids:string[]=[];
      let startIndex:number= batchIndex * 100;
      let endIndex:number = (batchIndex+1) * 100 < this.recycleBinItems.length? (batchIndex+1) * 100 : this.recycleBinItems.length;
      for(var index=startIndex; index < endIndex; index++)
      {
        ids.push(this.recycleBinItems[index].Id);
      }
      this.setState({
        message:`Restoring item from ${startIndex + 1} to ${endIndex}, please wait ...`,         
        messageType:MessageBarType.info
        });  

      let restoreRes = await RestAPIHelper.RestoreByIds(this.props.spHttpClient, this.state.affectedSite, ids);
      if(restoreRes)
      {
        this.setState({
          message:`Restored items from ${startIndex + 1} to ${endIndex}.`,         
          messageType:MessageBarType.success
          });  
      }
      else
      {
        this.setState({
          message:`Restoring item from ${startIndex} to ${endIndex} failed, please contact support.`,         
          messageType:MessageBarType.error
          }); 
          break;
      }
    }
  }

  private DoExport():void
  {
      // Export filtered recycle bin items
      SPOQAHelper.JSONToCSVConvertor(this.recycleBinItems, true, "RecycleBinItems");
  }
}