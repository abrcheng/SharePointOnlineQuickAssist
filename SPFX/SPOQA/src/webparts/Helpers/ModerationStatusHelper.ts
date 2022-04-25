import {SPHttpClient,ISPHttpClientOptions} from '@microsoft/sp-http';
import SPOQAHelper from './SPOQAHelper';
import RestAPIHelper from './RestAPIHelper';

// https://docs.microsoft.com/en-us/dotnet/api/microsoft.sharepoint.spmoderationstatustype?view=sharepoint-server
export enum SPModerationStatusType {
    Approved=0,	
    Denied=1,
    Pending=2,	
    Draft=3,	
    Scheduled=4
  }

  export enum ItemType 
  {
      ListItem=0,
      Document=1,
      Folder=2,
      UnKnown=-1
  }

  // Class for item which haven't been approved 
export class NotApprovedItem
{
    // name of the item, e.g. "Document.docx"
    public name:string;

    // full path of the item, e.g. https://chengc.sharepoint.com/sites/SPOQA/SubSite/Lists/SPOQAList/DispForm.aspx?ID=xx
    public url:string;    

    // parent folder of the item e.g. https://chengc.sharepoint.com/sites/SPOQA/SubSite/Shared Documents/Test 1/Level2/Level3
    public parentUrl:string;

    // ModerationStatus of the item, e.g. Pending=2
    public status:string;

    public constructor(name:string,url:string,parentUrl:string,status:string){
        this.name = name;
        this.url = url;
        this.parentUrl = parentUrl;
        this.status = status;
        }
}

// Class for getting the result of items which haven't been approved
export class NotApprovedItemsResult
{
    // Indicates whether execute the GetNotApprovedItems succesful or not
    public success:boolean;

    // If success=false, then detail error message should be set in this field 
    public error:string;

    // If there is any items include it parent folders haven't been approved, then push them into this array and return back
    public items:NotApprovedItem[];
    public itemType:ItemType;
    public constructor()
    {
        this.success = false;
        this.error ="";
        this.items = [];
        this.itemType = ItemType.UnKnown;
    }

}

// This class is used for operations about ModerationStatus 
// Check the approve status, OData__ModerationStatus
// All opertions need to consider about parent folders 
export class ModerationStatusHelper
{
    // List Item's URL format is https://chengc.sharepoint.com/sites/SPOQA/SubSite/Lists/SPOQAList/DispForm.aspx?ID=xx
    // Document's URL format is https://chengc.sharepoint.com/sites/SPOQA/SubSite/Shared Documents/xxxx/xxxx.xxx
    // Folder's URL format is https://chengc.sharepoint.com/sites/SPOQA/SubSite/Shared Documents/Test%201/Level2/Level3
    public static async GetNotApprovedItems(spHttpClient:SPHttpClient, siteAbsoluteUrl:string, listRootFolder:string, fullDocmentPath:string)
    {
        let res:NotApprovedItemsResult = new NotApprovedItemsResult();
        siteAbsoluteUrl = decodeURI(siteAbsoluteUrl);
        listRootFolder = decodeURI(listRootFolder);
        fullDocmentPath = decodeURI(fullDocmentPath);
        var listRootFolderFullUrl =  `https://${document.location.hostname}${listRootFolder}`;
        res.success= false;
        var apiUrl = `${siteAbsoluteUrl}/_api/web/`;
        if(fullDocmentPath.indexOf(listRootFolderFullUrl+"/")==-1) // fullDocmentPath must match listRootFolderFullUrl
        {           
            res.error = `The document full path ${fullDocmentPath} doesn't match list root folder ${listRootFolderFullUrl}.`;
            return res;
        }
        
        let itemType:ItemType=ItemType.UnKnown;
        if(fullDocmentPath.indexOf(".aspx?")>0) // the full document path should be pointted to list item
        {
          try
          {
            var urlParas = SPOQAHelper.ParseQueryString((fullDocmentPath.split(".aspx?"))[1]);   
            var itemId = urlParas["Id"]||urlParas["ID"];

            // https://chengc.sharepoint.com/sites/spoqa/_api/web/GetListUsingPath(DecodedUrl='/sites/SPOQA/LibAbc')
            apiUrl+=`GetListUsingPath(DecodedUrl='${listRootFolder}')/items(${itemId})`;
            var isItemExisting = await RestAPIHelper.IsObjectExisting(spHttpClient, apiUrl);
            if(!isItemExisting)
            {
                throw new Error(`Failed to load item via ${apiUrl}`);
            }
            itemType = ItemType.ListItem;
          }
          catch
          {
            res.error = `The document full path ${fullDocmentPath} is not a valid list item URL or the item doesn't exist.`;
            return res;
          }
        }
        else // check document (GetFileByUrl) or folder (GetFolderByServerRelativeUrl)
        {
            var relativeDocPath = fullDocmentPath.replace(`https://${document.location.hostname}`, "");
            apiUrl += `GetFileByUrl('${relativeDocPath}')`; 
            var isDocumentExisting = await RestAPIHelper.IsObjectExisting(spHttpClient, apiUrl);
            if(!isDocumentExisting)
            {
                // Try to check folder
                apiUrl = apiUrl.replace("/GetFileByUrl(", "/GetFolderByServerRelativeUrl(");
                var isFolderExisting = await RestAPIHelper.IsObjectExisting(spHttpClient, apiUrl);
                if(isFolderExisting)
                {
                    itemType = ItemType.Folder; 
                    apiUrl +="/ListItemAllFields";
                }
            }
            else
            {
                itemType = ItemType.Document;
                apiUrl +="/ListItemAllFields";
            }
        }
        
        if(itemType == ItemType.UnKnown)
        {
            res.error = `The full path ${fullDocmentPath} is not a valid URL or the item doesn't exist, ItemType is UnKnown.`;
            return res;
        }    
        
        res.itemType = itemType;

        // Get parent folder by REST API https://chengc.sharepoint.com/sites/SPOQA/SubSite/_api/Lists/getByTitle('SPOQAList')/items(4)?$select=*,FileDirRef
        var parentFolder = "";
        while(parentFolder!=listRootFolder)
        {
           var itemInfoRes:any = await ModerationStatusHelper.GetItemModerationStatusAndFileDirRef(spHttpClient, apiUrl);
           if(!itemInfoRes.success)
           {
                res.error = itemInfoRes.error;
                return res;
           }

           var itemInfo:any = itemInfoRes.res;
           if(itemInfo.OData__ModerationStatus && itemInfo.OData__ModerationStatus >0)
           {   
               var itemName = itemInfo.Title;
               if(!itemInfo.Title)
               {
                    itemName = itemInfo.FileRef.replace(itemInfo.FileDirRef,"").replace("/","");
               }
                
               var absParentFolder = `https://${document.location.hostname}${itemInfo.FileDirRef}`;
               var absItemUrl = fullDocmentPath;
               if(parentFolder)
               {
                    absItemUrl =  `https://${document.location.hostname}${itemInfo.FileRef}`;
               }
               var notApprovedItem = new NotApprovedItem(itemName, absItemUrl, absParentFolder,SPModerationStatusType[itemInfo.OData__ModerationStatus]);
               res.items.push(notApprovedItem);
           }

           parentFolder = itemInfo.FileDirRef;
           apiUrl = `${siteAbsoluteUrl}/_api/web/GetFolderByServerRelativeUrl('${parentFolder}')/ListItemAllFields`;
        }
        
        res.success= true;
        return res;
    }

    private static async GetItemModerationStatusAndFileDirRef(spHttpClient:SPHttpClient, itemApiUrl:string)
    {
        itemApiUrl = itemApiUrl +"?$select=*,FileDirRef,FileRef";
        var itemRes = await spHttpClient.get(itemApiUrl, SPHttpClient.configurations.v1);
        if(itemRes.ok)
        {
            var resJson = await itemRes.json();
            if(resJson.error)
            {
                console.error(`${resJson.error.message} ${itemApiUrl}.`);
                return {success:false, error:resJson.error.message};               
            }

            return {success:true,res:resJson};
        }
        
        var failedMsg = `Failed to get data from request ${itemApiUrl}.`;
        console.error(failedMsg);
        return {success:false, error:failedMsg};   
    }
}
