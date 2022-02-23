import * as React from 'react';
import {  
    PrimaryButton,    
    TextField,
    Label,
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
import { forEach } from 'lodash';
export default class RestoreItemsQA extends React.Component<ISharePointOnlineQuickAssistProps>
{
  public recycleBinItems:IRestoreItem[];
  public state = {
    deleteByUser: this.props.currentUser.loginName,
    deleteStartDate:null,
    deleteEndDate:null,
    pathFilter:"",    
    affectedSite:this.props.webAbsoluteUrl,
    queried:false,
    currentItems:null        
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
          <div id="RestoreItemsQA_QueryResultSection">
              <RestoreItemsQAGrid items={this.state.currentItems}/>
          </div>
          <div id="RestoreItemsQA_CommandButtonsSection">
              <PrimaryButton
                        text="Query Items"
                        style={{ display: 'inline', marginTop: '10px' }}
                        onClick={() => {this.QueryRecycleBinItems();}}
                        />
              {this.state.queried && false?
                        <PrimaryButton
                            text="Restore"
                            style={{ display: 'inline', marginTop: '10px', marginLeft:"10px"}}
                            // onClick={() => {this.ShowRemedySteps();}}
                        />:
                        null}
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
         while(currentCount ==500 || currentCount==-1)
         {
           var recycleItems = await RestAPIHelper.Getrecyclebinitems(this.props.spHttpClient, this.state.affectedSite,pageInfo,500,false, 1, 3);

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
              if(this.IsMatchFilter(currentItem))
              {
                this.recycleBinItems.push(currentItem);
              }
              lastId = currentItem.Id;
              lastTitle = currentItem.Title;
              lastDeletedDate = currentItem.DeletedDate;
           }
            
           console.log(`RestAPIHelper.Getrecyclebinitems with page info ${pageInfo}`);
           pageInfo = URI_Encoding.encodeURIComponent(`id=${lastId}&title=${lastTitle}&searchValue=${lastDeletedDate}`);  
           SPOQASpinner.Show(`Queried ${this.recycleBinItems.length} ......`);      
         }

         this.setState({currentItems:this.recycleBinItems});
         // this.setState({queried:true});
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
      matched = matched&&(this.state.pathFilter.toLowerCase() == item.DirName.toLowerCase());
    }

    if(this.state.deleteStartDate)
    {
        matched = matched&&(this.state.deleteStartDate < item.DeletedDate);
    }
    
    if(this.state.deleteEndDate)
    {
        matched = matched&&(this.state.deleteEndDate > item.DeletedDate);
    }

    return matched;
  }

}