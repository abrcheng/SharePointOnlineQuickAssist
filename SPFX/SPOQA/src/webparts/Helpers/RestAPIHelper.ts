import {SPHttpClient,ISPHttpClientOptions} from '@microsoft/sp-http';
import { format } from 'office-ui-fabric-react';
import SPOQAHelper from './SPOQAHelper';

export default class RestAPIHelper
{
    public static async GetUserInfoFromUserProfile(user: string, spHttpClient:SPHttpClient, webAbsoluteUrl:string)
    {      
        var apiUrl = `${webAbsoluteUrl}/_api/SP.UserProfiles.PeopleManager/GetPropertiesFor(accountName=@v)?@v='i:0%23.f|membership|${user}'`;
        var res = await RestAPIHelper.CallGetRest(apiUrl, spHttpClient);
        return await res;
    }
    
    public static async GetUserFromUserInfoList(user:string, spHttpClient:SPHttpClient, webAbsoluteUrl:string)
    {
      var account =  `i:0#.f|membership|${user}`;
      var apiUrl = `${webAbsoluteUrl}/_api/web/SiteUserInfoList/items?$filter=Name eq '${encodeURIComponent(account)}'`;      
      var res = await RestAPIHelper.CallGetRest(apiUrl, spHttpClient);
      if(res.value.length>=1)
      {
        return await res.value[0];
      }
      else
      {
         console.log(`Result is empty for the API URL ${apiUrl}`);
         return null;
      }
    }
    
    // SP.Data.UserInfoItem
    public static async FixJobTitleInUserInfoList(userId:number, spHttpClient:SPHttpClient, webAbsoluteUrl:string, newJobTitle:string, successCallBack:Function, failedCallback:Function) 
    {    
      const context: SP.ClientContext = new SP.ClientContext(webAbsoluteUrl);     
      const userItem: SP.ListItem = context.get_web().get_siteUserInfoList().getItemById(userId); 
      userItem.set_item('JobTitle', newJobTitle);      
      userItem.update();
      context.executeQueryAsync((sender: any, args: SP.ClientRequestSucceededEventArgs): void => {successCallBack();}, (sender: any, args: SP.ClientRequestSucceededEventArgs): void => {failedCallback();});      
    }    
    
    public static TestConnectMySite(spHttpClient:SPHttpClient, mySiteHost:string)
    {
      const context: SP.ClientContext = new SP.ClientContext(mySiteHost);  
      let userPhotoLib:SP.List = context.get_web().get_lists().getByTitle("User Photos");
      context.load(userPhotoLib);
      context.executeQueryAsync(
        (sender: any, args: SP.ClientRequestSucceededEventArgs): void => {console.log(userPhotoLib.get_title()+" loaded in TestConnectMySite");}, 
        (sender: any, args: SP.ClientRequestSucceededEventArgs): void => {console.log("Failed to run TestConnectMySite");});
    }

    public static async FixJobTitleInUserProfile(user:string,spHttpClient:SPHttpClient, webAbsoluteUrl:string, newJobTitle:string)
    {
        let apiUrl = webAbsoluteUrl + "/_api/SP.UserProfiles.PeopleManager/SetSingleValueProfileProperty";  
        let userData = {  
            'accountName': "i:0#.f|membership|" + user,  
            'propertyName': "SPS-JobTitle", 
            'propertyValue': newJobTitle  
        };
      
        let spOpts = {  
            headers: {  
                'Accept': 'application/json;odata=nometadata',  
                'Content-type': 'application/json;odata=verbose',  
                'odata-version': '',  
            },  
            body: JSON.stringify(userData)  
        };  
        var res = await spHttpClient.post(apiUrl, SPHttpClient.configurations.v1, spOpts); 
        if(res.ok)
        {
          var responseJson = await res.json();
          console.log(`FixJobTitleInUserProfile done for API url ${apiUrl}`);          
          return await responseJson;
        }       
    }

    public static async GetUserInfoFromSite(user:string, spHttpClient:SPHttpClient, webAbsoluteUrl:string)
    { 
      var account =  `i:0#.f|membership|${user}`;
      var apiUrl = `${webAbsoluteUrl}/_api/web/siteusers(@v)?@v='${encodeURIComponent(account)}'`;
      var res = await RestAPIHelper.CallGetRest(apiUrl, spHttpClient);
      return res;
    }

    private static async CallGetRest(apiUrl:string, spHttpClient:SPHttpClient, )
    {
      var res = await spHttpClient.get(apiUrl, SPHttpClient.configurations.v1);
      if(res.ok)
      {
        var responseJson = await res.json();
        console.log(`GetUserInfoFromUserProfile done for API url ${apiUrl}`);          
        return await responseJson;
      }
      else
      {
        var message = `Failed GetUserInfoFromSite for API url ${apiUrl}`;
        console.log(message);
        Promise.reject(message);
      }
    }

    public static async GetQueryUser(url:string, spHttpClient:SPHttpClient) {
      
    }

    public static async GetSerchResults(spHttpClient:SPHttpClient, siteAbsoluteUrl:string, targetwebAbsoluteUrl:string, contentClass:string)
    { 
      var contentClassStr = `*`;
      if(contentClass == "Site")
      {
        contentClassStr = `ContentClass:STS_Site Path:"${targetwebAbsoluteUrl}"`;
      }
      var apiUrl = `${siteAbsoluteUrl}/search/_api/search/query?querytext='${contentClassStr}'&SelectProperties='Path,Title'&rowlimit=10`;

      var res = await spHttpClient.get(apiUrl, SPHttpClient.configurations.v1);
      if(res.ok)
      {
        var responseJson = await res.json();
        console.log(`GetSerchResults done for API url ${apiUrl}`);          
        return await responseJson;
      }
      else
      {
        var message = `Failed GetSerchResults for API url ${apiUrl}`;
        console.log(message);
        Promise.reject(message);
      }
    }

    public static async GetWeb(spHttpClient:SPHttpClient, webAbsoluteUrl:string)
    { 
      var apiUrl = `${webAbsoluteUrl}/_api/web`;

      var res = await spHttpClient.get(apiUrl, SPHttpClient.configurations.v1);
      if(res.status == 404)
      {
        return false;
      }
      return true;
    }

    public static async GetLists(spHttpClient:SPHttpClient, siteAbsoluteUrl:string)
    {
        console.log(`Start to load list for the site ${siteAbsoluteUrl}`);
        var apiUrl = `${siteAbsoluteUrl}/_api/web/Lists?$select=Title,BaseTemplate,BaseType&rowlimit=5000`;
        var res = await spHttpClient.get(apiUrl, SPHttpClient.configurations.v1);
        var resJson = await res.json();
        return resJson.value;
    }

    public static async SearchDocumentByFullPath(spHttpClient:SPHttpClient, siteAbsoluteUrl:string, fullPath:string)
    {
        var queryText = `Path:"${fullPath}"`;
        var apiUrl = `${siteAbsoluteUrl}/_api/search/query?querytext='${queryText}'&SelectProperties='Path,Title'`;
        var res = await spHttpClient.get(apiUrl, SPHttpClient.configurations.v1);
        var resJson = await res.json();
        return resJson.PrimaryQueryResult.RelevantResults;
    }

    public static async IsWebNoCrawl(spHttpClient:SPHttpClient, siteAbsoluteUrl:string)
    {
        var apiUrl = `${siteAbsoluteUrl}/_api/web`;
        var res = await spHttpClient.get(apiUrl, SPHttpClient.configurations.v1);
        var resJson = await res.json();
        return resJson.NoCrawl;
    }

    public static async IsListNoCrawl(spHttpClient:SPHttpClient, siteAbsoluteUrl:string, listTitle:string)
    {
      var apiUrl = `${siteAbsoluteUrl}/_api/web/Lists/getByTitle('${listTitle}')`;
      var res = await spHttpClient.get(apiUrl, SPHttpClient.configurations.v1);
      var resJson = await res.json();
      return resJson.NoCrawl;
    }

    public static async IsListMissDisplayForm(spHttpClient:SPHttpClient, siteAbsoluteUrl:string, listTitle:string)
    {
      var apiUrl = `${siteAbsoluteUrl}/_api/web/Lists/getByTitle('${listTitle}')/Forms`;
      var res = await spHttpClient.get(apiUrl, SPHttpClient.configurations.v1);
      var resJson = await res.json();
      let hasDisplayForm:boolean = false;
      let displayFormUrl = "";
      resJson.value.forEach(form => {
        if(form.FormType == 4)
        {
          if(form.ServerRelativeUrl)
          {
            displayFormUrl = form.ServerRelativeUrl;            
          }
        }
      });

      if(displayFormUrl && displayFormUrl !="")
      {
        try
        {
          apiUrl = `${siteAbsoluteUrl}/_api/web/GetFileByUrl('${displayFormUrl}')`;
          res = await spHttpClient.get(apiUrl, SPHttpClient.configurations.v1);
          resJson = await res.json();
          hasDisplayForm = true;
        }
        catch(err)
        {
          console.log(err);
        }
      }

      return !hasDisplayForm;
    }

    public static async IsDocumentInDraftVersion(spHttpClient:SPHttpClient, siteAbsoluteUrl:string, isDocument:boolean, listTitle:string, fullDocmentPath:string)
    {
        // https://chengc.sharepoint.com/_api/web/Lists/getByTitle('TestMMList')/items(1)
        // https://chengc.sharepoint.com/_api/web/GetFileByUrl('/test123/Document.docx')/ListItemAllFields
        var apiUrl = `${siteAbsoluteUrl}/_api/web/`;
        if(isDocument)
        {
           var relativeDocPath = fullDocmentPath.replace(`https://${document.location.hostname}`, "");
           apiUrl += `GetFileByUrl('${relativeDocPath}')/ListItemAllFields`; 
        }
        else
        {       
          var urlParas = SPOQAHelper.ParseQueryString((fullDocmentPath.split(".aspx?"))[1]);   
          var itemId = urlParas["Id"]||urlParas["ID"];
          apiUrl+=`Lists/getByTitle('${listTitle}')/items(${itemId})`;
        }
        console.log(`Will call API ${apiUrl}`);
        var res = await spHttpClient.get(apiUrl, SPHttpClient.configurations.v1);
        var resJson = await res.json();
        var versionStr = resJson.OData__UIVersionString;
        var minVersion = (versionStr.split("."))[1];
        return minVersion > '0';
    }    
}