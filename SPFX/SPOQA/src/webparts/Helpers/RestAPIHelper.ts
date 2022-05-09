import {SPHttpClient,ISPHttpClientOptions} from '@microsoft/sp-http';
import SPOQAHelper from './SPOQAHelper';
import { WebPartContext } from '@microsoft/sp-webpart-base';

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

    public static async GetGroupFromUserInfoList(groupId:string, spHttpClient:SPHttpClient, webAbsoluteUrl:string)
    {      
      var apiUrl = `${webAbsoluteUrl}/_api/web/SiteUserInfoList/items?$filter=substringof('${groupId}',Name)'`;      
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
    
    public static async FixUserInfoItem(userId:number, spHttpClient:SPHttpClient, webAbsoluteUrl:string,properties:any ,successCallBack:Function, failedCallback:Function)
    {     
      const context: SP.ClientContext = new SP.ClientContext(webAbsoluteUrl);     
      const userItem: SP.ListItem = context.get_web().get_siteUserInfoList().getItemById(userId); 
      properties.forEach(p => {
        userItem.set_item(p.key, p.value);    
      });       
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

    public static async GetUserReadPermissions(user:string, spHttpClient:SPHttpClient, webAbsoluteUrl:string)
    {
      var account =  `i:0#.f|membership|${user}`;
      var apiUrl = `${webAbsoluteUrl}/_api/web/getUserEffectivePermissions(@username)?@username='${encodeURIComponent(account)}'`;      
      var res = await spHttpClient.get(apiUrl, SPHttpClient.configurations.v1);
      if(res.ok)
      {
        var responseJson = await res.json();
        console.log(`GetUserPermissions done for API url ${apiUrl}`);
        var permissions = new SP.BasePermissions();
        permissions.fromJson(responseJson);
        //let result = {};
        /*
        for(var levelName in SP.PermissionKind) {
            if (SP.PermissionKind.hasOwnProperty(levelName)) {
                 let permLevel:SP.PermissionKind = levelName ;
                if(permissions.has(permLevel))
                {
                  result[levelName] = true;
                }
                else
                {
                  result[levelName] = false;
                }
            }     
        }*/
        if((permissions.has(SP.PermissionKind.viewListItems)) && (permissions.has(SP.PermissionKind.openItems)) && (permissions.has(SP.PermissionKind.viewVersions)) && (permissions.has(SP.PermissionKind.viewFormPages)) && (permissions.has(SP.PermissionKind.open))
        && (permissions.has(SP.PermissionKind.viewPages)) && (permissions.has(SP.PermissionKind.browseUserInfo)) && (permissions.has(SP.PermissionKind.useClientIntegration)) && (permissions.has(SP.PermissionKind.useRemoteAPIs)) && (permissions.has(SP.PermissionKind.createAlerts)))
        {
          return true;
        }
        else
        {
          return false;
        }
      }
      else
      {
        var message = `Failed GetUserReadPermissions for API url ${apiUrl}`;
        console.log(message);
        Promise.reject(message);
      }
    }

    public static async GetSiteGroupId(spHttpClient:SPHttpClient, ctx:WebPartContext, siteAbsoluteUrl:string)
    {
      // var siteURL = ctx.pageContext.site.absoluteUrl;
      var apiUrl = `${siteAbsoluteUrl}/_api/site`;      
      var res = await spHttpClient.get(apiUrl, SPHttpClient.configurations.v1);
      if(res.ok)
      {
        var responseJson = await res.json();
        console.log(`GetSiteGroupId done for API url ${apiUrl}`);
        return await responseJson['GroupId'];
      }
      else
      {
        var message = `Failed GetSiteGroupId for API url ${apiUrl}`;
        console.log(message);
        Promise.reject(message);
      }
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
        contentClassStr = `ContentClass:STS_Site Path="${targetwebAbsoluteUrl}"`;
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

    public static async GetSiteId(spHttpClient:SPHttpClient, siteAbsoluteUrl:string)
    {
      var apiUrl = `${siteAbsoluteUrl}/_api/site/ID`;      
      var res = await spHttpClient.get(apiUrl, SPHttpClient.configurations.v1);
      if(res.ok)
      {
        var responseJson = await res.json();
        console.log(responseJson);
        console.log(`GetSiteId done for API url ${apiUrl}`);
        return await responseJson.value;
      }
      else
      {
        var message = `Failed GetSiteId for API url ${apiUrl}`;
        console.log(message);
        Promise.reject(message);
        return null;
      }
    }

    public static async GetLists(spHttpClient:SPHttpClient, siteAbsoluteUrl:string)
    {
        console.log(`Start to load list for the site ${siteAbsoluteUrl}`);
        var apiUrl = `${siteAbsoluteUrl}/_api/web/Lists?$select=Id,Title,BaseTemplate,BaseType,RootFolder/ServerRelativeUrl,WriteSecurity,ReadSecurity&$expand=RootFolder&rowlimit=5000`;
        var res = await spHttpClient.get(apiUrl, SPHttpClient.configurations.v1);
        var resJson = await res.json();
        return resJson.value;
    }

    public static async GetPageByUrl(spHttpClient:SPHttpClient, pageUrl:string)
    {
      var res = await spHttpClient.get(pageUrl, SPHttpClient.configurations.v1);
      if(res.ok)
      {
        var html = await res.text();
        const parser = new DOMParser();
        const parsedDocument = parser.parseFromString(html, "text/html");
        return parsedDocument;        
      }
      else
      {
        console.log(`Failed to load ${pageUrl}`);
        return null;
      }
    }
    
    public static async GetWikiField(spHttpClient:SPHttpClient, webAbsoluteUrl:string, pageListId:string, pageItemId:string)
    {
      var requestUrl = webAbsoluteUrl + "/_api/web/lists/GetById('" + pageListId + "')/items/GetById('" + pageItemId + "')";
      var res = await spHttpClient.get(requestUrl, SPHttpClient.configurations.v1);
      if(res.ok)
      {
        var responseJson = await res.json();
        const parser = new DOMParser();
        const parsedDocument = parser.parseFromString(responseJson.WikiField, "text/html");
        return parsedDocument;      
      }
      else
      {
        console.log(`Failed to get wikiField, status code: ${res.status}`);
        return null;
      }
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

    public static async IsListMissNewForm(spHttpClient:SPHttpClient, siteAbsoluteUrl:string, listTitle:string)
    {
      var apiUrl = `${siteAbsoluteUrl}/_api/web/Lists/getByTitle('${listTitle}')/Forms`;
      var res = await spHttpClient.get(apiUrl, SPHttpClient.configurations.v1);
      var resJson = await res.json();
      let hasNewForm:boolean = false;
      let newFormUrl = "";
      resJson.value.forEach(form => {
        if(form.FormType == 8)
        {
          if(form.ServerRelativeUrl)
          {
            newFormUrl = form.ServerRelativeUrl;            
          }
        }
      });

      if(newFormUrl && newFormUrl !="")
      {
        try
        {
          apiUrl = `${siteAbsoluteUrl}/_api/web/GetFileByUrl('${newFormUrl}')`;
          res = await spHttpClient.get(apiUrl, SPHttpClient.configurations.v1);
          resJson = await res.json();
          hasNewForm = true;
        }
        catch(err)
        {
          console.log(err);
        }
      }

      return !hasNewForm;
    }

    public static async IsListMissEditForm(spHttpClient:SPHttpClient, siteAbsoluteUrl:string, listTitle:string)
    {
      var apiUrl = `${siteAbsoluteUrl}/_api/web/Lists/getByTitle('${listTitle}')/Forms`;
      var res = await spHttpClient.get(apiUrl, SPHttpClient.configurations.v1);
      var resJson = await res.json();
      let hasEditForm:boolean = false;
      let editFormUrl = "";
      resJson.value.forEach(form => {
        if(form.FormType == 6)
        {
          if(form.ServerRelativeUrl)
          {
            editFormUrl = form.ServerRelativeUrl;            
          }
        }
      });

      if(editFormUrl && editFormUrl !="")
      {
        try
        {
          apiUrl = `${siteAbsoluteUrl}/_api/web/GetFileByUrl('${editFormUrl}')`;
          res = await spHttpClient.get(apiUrl, SPHttpClient.configurations.v1);
          resJson = await res.json();
          hasEditForm = true;
        }
        catch(err)
        {
          console.log(err);
        }
      }

      return !hasEditForm;
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
        if(versionStr)
        {
          console.log(`UIVersionString is ${versionStr} for the document ${fullDocmentPath}`);
          var minVersion = (versionStr.split("."))[1];
          return minVersion > '0';
        }
        else // version haven't been enabled
        {
          console.log(`Version haven't enabled for library ${listTitle}`);
          return false;
        }
    } 
    
    public static async FixListNoCrawl(spHttpClient:SPHttpClient, siteAbsoluteUrl:string, listTitle:string)
    {      
      let apiUrl:string = `${siteAbsoluteUrl}/_api/web/Lists/getByTitle('${listTitle}')`;  
      let listData = { 
          __metadata:
          {
              type: "SP.List"
          },       
          NoCrawl: false
        };
    
      let spOpts = {  
            headers: {              
              "Accept": "application/json;odata=verbose",            
              "Content-Type": "application/json;odata=verbose",            
              "IF-MATCH": "*",            
              "X-HTTP-Method": "MERGE",
              "odata-version": "3.0"               
            },  
            body: JSON.stringify(listData)  
        };  
      var res = await spHttpClient.post(apiUrl, SPHttpClient.configurations.v1, spOpts);      
      return res;
    }
  
    public static async FixWebNoCrawl(spHttpClient:SPHttpClient, siteAbsoluteUrl:string)
    {
      let apiUrl = `${siteAbsoluteUrl}/_api/web`;
      let listData = { 
          __metadata:
          {
              type: "SP.Web"
          },       
          NoCrawl: false
        };
    
      let spOpts = {  
            headers: {              
              "Accept": "application/json;odata=verbose",            
              "Content-Type": "application/json;odata=verbose",            
              "IF-MATCH": "*",            
              "X-HTTP-Method": "MERGE",
              "odata-version": "3.0"            
            },  
            body: JSON.stringify(listData)  
        };  
      var res = await spHttpClient.post(apiUrl, SPHttpClient.configurations.v1, spOpts);      
      return res;
    }

    public static async FixDraftVersion(spHttpClient:SPHttpClient, siteAbsoluteUrl:string, isDocument:boolean, listTitle:string, fullDocmentPath:string)
    {
        // Only document will have the draft version 
        if(!isDocument)
        {
          console.log("Only document will have the draft version, ignore fix request for isDocument===false");
        }
        
        // /_api/web/getfilebyserverrelativeurl('Server Relative URL%')/CheckIn(comment='Check-in by SharePointOnlineQuickAssist',checkintype=1)
        // "X-HTTP-Method": "PATCH",
        // https://chengc.sharepoint.com/sites/abc/TestSPOQA/_api/web/GetFileByServerRelativePath(DecodedUrl=@a1)/Publish(@a2)?@a1=%27%2Fsites%2Fabc%2FTestSPOQA%2FShared%20Documents%2FDocument2%2Edocx%27&@a2=%27Looks%20good%27
        var resJson;
        var relativeDocPath = fullDocmentPath.replace(`https://${document.location.hostname}`, "");
        let spOpts = {  
          headers: {              
            "Accept": "application/json;odata=verbose",            
            "Content-Type": "application/json;odata=verbose",            
            "IF-MATCH": "*",            
            "X-HTTP-Method": "PATCH"            
          }          
        };  
                
          let apiUrl:string = `${siteAbsoluteUrl}/_api/web/GetFileByUrl('${relativeDocPath}')/CheckIn(comment='Check-in by SharePointOnlineQuickAssist',checkintype=1)`;           
          var res = await spHttpClient.post(apiUrl, SPHttpClient.configurations.v1, spOpts); 
          resJson = await res.json();
          console.log(resJson);
          if(resJson.error)
          {
             apiUrl = `${siteAbsoluteUrl}/_api/web/GetFileByUrl('${relativeDocPath}')/Publish('Published by SharePointOnlineQuickAssist')`;           
             res = await spHttpClient.post(apiUrl, SPHttpClient.configurations.v1, spOpts);   
            console.log(`${apiUrl} OK? ${res.ok}`);
          }
     

      return resJson;
    }
    
    // Return list info which returned by https://xxxxx.sharepoint.com/sites/abc/_api/web/Lists/getByTitle('xxxxx')
    // Get ExcludeFromOfflineClient of list https://chengc.sharepoint.com/sites/abc/_api/web/Lists/getByTitle('GifLib')?$select=ExcludeFromOfflineClient
    // properties: ExcludeFromOfflineClient,ForceCheckout,DraftVersionVisibility,EnableModeration,ValidationFormula,ValidationMessage
    public static async GetListInfo(spHttpClient:SPHttpClient, siteAbsoluteUrl:string, listTitle:string, properties:string[])
    {     
      let selectStr:string = RestAPIHelper.BuildSelectStr(properties);
      var apiUrl = `${siteAbsoluteUrl}/_api/web/Lists/getByTitle('${listTitle}')${selectStr}`;
      var res = await spHttpClient.get(apiUrl, SPHttpClient.configurations.v1);
      var resJson = await res.json();
      return resJson;
    }
    
    // https://chengc.sharepoint.com/sites/abc/_api/web/Lists/getByTitle('GifLib')/SchemaXml
    public static async GetListFields(spHttpClient:SPHttpClient, siteAbsoluteUrl:string, listTitle:string)
    {
      var apiUrl = `${siteAbsoluteUrl}/_api/web/Lists/getByTitle('${listTitle}')/Fields`;
      var res = await spHttpClient.get(apiUrl, SPHttpClient.configurations.v1);
      var resJson = await res.json();
      return resJson.value;
    }
    
    // https://chengc.sharepoint.com/sites/abc/_api/site/Features
    // Check if the feature (e.g.  7c637b23-06c4-472d-9a9a-7c175762c5c4) is enabled or not in the site collection
    public static async IsSiteFeatureEnabled(spHttpClient:SPHttpClient, siteAbsoluteUrl:string, featureId:string)
    {
      var apiUrl = `${siteAbsoluteUrl}/_api/site/Features`;
      var res = await spHttpClient.get(apiUrl, SPHttpClient.configurations.v1);
      var resJson = await res.json();
      let enabled = false;
      for(var i=0; i<resJson.value.length;i++)
      {
        if(resJson.value[i].DefinitionId.toLowerCase() == featureId.toLowerCase())
        {
          enabled = true;
          break;
        }
      }

      return enabled;
    }    

    // https://chengc.sharepoint.com/sites/abc/TestSPOQA/_api/web?$select=ExcludeFromOfflineClient
    public static async GetWebInfo(spHttpClient:SPHttpClient, siteAbsoluteUrl:string, properties:string[])
    {
      let selectStr:string = RestAPIHelper.BuildSelectStr(properties);
      var apiUrl = `${siteAbsoluteUrl}/_api/web${selectStr}`;
      var res = await spHttpClient.get(apiUrl, SPHttpClient.configurations.v1);
      var resJson = await res.json();
      return resJson;
    }
    
    // https://chengc.sharepoint.com/sites/abc/TestSPOQA/_api/web/ParentWeb
    public static async GetParentWebUrl(spHttpClient:SPHttpClient, siteAbsoluteUrl:string)
    {
      let parentWebUrl = "";
      var apiUrl = `${siteAbsoluteUrl}/_api/web/ParentWeb`;
      var res = await spHttpClient.get(apiUrl, SPHttpClient.configurations.v1);
      var resJson = await res.json();
      if(resJson.ServerRelativeUrl)
      {
        let url:URL = new URL(siteAbsoluteUrl);
        let rootSiteUrl = `${url.protocol}//${url.hostname}`;
        parentWebUrl = `${rootSiteUrl}${resJson.ServerRelativeUrl}`;
      }

      return parentWebUrl;
    }

    public static async GetWebExcludeFromOfflineClient(spHttpClient:SPHttpClient, siteAbsoluteUrl:string)
    {
        let properties:string[] = ["ExcludeFromOfflineClient"];
        let resList:any[] = [];
        var hasParentWeb = true;
        let currentWebUrl = siteAbsoluteUrl;
        while(hasParentWeb)
        {
            var webInfo = await RestAPIHelper.GetWebInfo(spHttpClient, currentWebUrl, properties);
            resList.push({webUrl:currentWebUrl,
               ExcludeFromOfflineClient:webInfo.ExcludeFromOfflineClient,
               RemedyUrl:`${currentWebUrl}/_layouts/15/srchvis.aspx`});
            currentWebUrl = await RestAPIHelper.GetParentWebUrl(spHttpClient, currentWebUrl);
            hasParentWeb = currentWebUrl && currentWebUrl!="";
        }

        return resList;
    }

    // https://chengc.sharepoint.com/sites/abc/_api/web/Lists/getByTitle('GifLib')/GetUserEffectivePermissions('i%3A0%23.f%7Cmembership%7Cjohnb%40chengc.onmicrosoft.com')
    // permission: SP.PermissionKind.editListItems
    public static async HasPermissionOnList(spHttpClient:SPHttpClient, siteAbsoluteUrl:string, listTitle:string, user:string, permission:SP.PermissionKind)
    {      
      var account =  `i:0#.f|membership|${user}`;
      var apiUrl = `${siteAbsoluteUrl}/_api/web/Lists/getByTitle('${listTitle}')/GetUserEffectivePermissions('${encodeURIComponent(account)}')`;      
      return await RestAPIHelper.HasPermssionOnOject(spHttpClient, apiUrl, permission);
    }
    
    public static async Getrecyclebinitems(spHttpClient:SPHttpClient, siteAbsoluteUrl:string, pageInfo:string, rowLimit:number,isAscending:boolean, itemState:number, orderby:number)
    {
       // https://chengc.sharepoint.com/sites/abc/_api/site/getrecyclebinitems?rowLimit='100'&isAscending=false&itemState=1&orderby=3
       // 'id=dbe08209-a916-4762-8390-200aeefe91f2&title=Table of Contents.docx&searchValue=2021-12-21T08:25:47' => encode => pagingInfo
       // https://chengc.sharepoint.com/sites/abc/_api/site/getrecyclebinitems?rowLimit=%27101%27&isAscending=false&itemState=1&orderby=3&pagingInfo=%27id%3Ddbe08209-a916-4762-8390-200aeefe91f2%26title%3DTable%20of%20Contents.docx%26searchValue%3D2021-12-21T08%3A25%3A47%27
       
       var apiUrl = `${siteAbsoluteUrl}/_api/site/getrecyclebinitems?rowLimit='${rowLimit}'&isAscending=${isAscending}&itemState=${itemState}&orderby=${orderby}`;
       if(pageInfo && pageInfo.length > 0)
       {
          apiUrl = `${apiUrl}&pagingInfo=${pageInfo}`;            
       }

       var res = await spHttpClient.get(apiUrl, SPHttpClient.configurations.v1);
       if(res.ok)
       {
          var resJson = await res.json();
          console.log(`Getrecyclebinitems done for API url ${apiUrl}`);          
          return resJson;
       }
       else
       {
        var message = `Failed Getrecyclebinitems for API url ${apiUrl}`;
        console.log(message);       
       }
    }

    public static async RestoreByIds(spHttpClient:SPHttpClient, siteAbsoluteUrl:string, ids:string[])
    {
        let requestBody:any={"ids":ids, "bRenameExistingItems":"true"};
        requestBody = JSON.stringify(requestBody);
        var apiUrl = `${siteAbsoluteUrl}/_api/site/RecycleBin/RestoreByIds`;
        let spOpts = {  
          headers: {              
            "Accept": "application/json;odata=verbose",            
            "Content-Type": "application/json;odata=verbose",            
            "IF-MATCH": "*"                 
          },
          body:requestBody
        };  

        var res = await spHttpClient.post(apiUrl, SPHttpClient.configurations.v1, spOpts);
        if(res.ok)
        {            
            console.log(`RestoreByIds done for API url ${apiUrl}`);          
            return {success:true};
        }
        else
        {
          var resJson = await res.json();
          var message = `Failed RestoreByIds for API url ${apiUrl}`;
          console.log(JSON.stringify(requestBody));    
          return  resJson;         
        }
    }
    
    // TODO check file existing or not GetFileByServerRelativeUrl 
    // https://chengc.sharepoint.com/sites/SPOQA/_api/web/GetFileByServerRelativeUrl('/sites/SPOQA/SitePages/Home.aspx')
    public static async IsDocumentExisting(spHttpClient:SPHttpClient, siteAbsoluteUrl:string, fileServerRelativeUrl:string)
    {
      var apiUrl = `${siteAbsoluteUrl}/_api/web/GetFileByServerRelativeUrl('${fileServerRelativeUrl}')`;
      var res = await spHttpClient.get(apiUrl, SPHttpClient.configurations.v1);
      if(res.ok)
      {
        console.log(`Got file for url ${fileServerRelativeUrl}`);
        return {success:true};
      }
      else
      {
        var resJson = await res.json();
        console.log(`Failed to get file for url ${fileServerRelativeUrl} by error message ${resJson.error.message}`);
        return {success:false, message:resJson.error.message};
      }
    }

    // TODO check user's pmerssion on document 
    // https://chengc.sharepoint.com/sites/SPOQA/_api/web/GetFileByServerRelativeUrl('/sites/SPOQA/SitePages/Home.aspx')/ListItemAllFields/GetUserEffectivePermissions('i%3A0%23.f%7Cmembership%7Cjohnb%40chengc.onmicrosoft.com')
    public static async HasPermissionOnDocument(spHttpClient:SPHttpClient, siteAbsoluteUrl:string, documentUrl:string, user:string, permission:SP.PermissionKind)
    {
      var account =  `i:0#.f|membership|${user}`;
      var apiUrl = `${siteAbsoluteUrl}/_api/web/GetFileByServerRelativeUrl('${documentUrl}')/ListItemAllFields/GetUserEffectivePermissions('${encodeURIComponent(account)}')`;  
      return await RestAPIHelper.HasPermssionOnOject(spHttpClient, apiUrl, permission); 
    }

    // TODO check file without check-in version 
    // https://chengc.sharepoint.com/sites/SPOQA/_layouts/15/ManageCheckedOutFiles.aspx?List=%7B6ED564A0%2DAA16%2D4AB9%2D8721%2D1AC6EC3F6354%7D    
    public static async HasFileWithOutCheckInVersion(spHttpClient:SPHttpClient, siteAbsoluteUrl:string,listId:string)
    {
      var apiUrl = `${siteAbsoluteUrl}/_layouts/15/ManageCheckedOutFiles.aspx?List={${listId}}`;
      var res = await spHttpClient.get(apiUrl, SPHttpClient.configurations.v1);
      if(res.ok)
      {
        var html = await res.text();
        const parser = new DOMParser();
        const parsedDocument = parser.parseFromString(html, "text/html");
        return parsedDocument.querySelectorAll("#onetidTable tr").length >=3;        
      }
      else
      {
        console.log(`Failed to load ${apiUrl}`);
        return {success:false};
      }
    }

    private static BuildSelectStr(properties:string[]):string
    {
      var selectStr="";
      if(properties && properties.length >0)
      {
        selectStr="?$select=";
        properties.forEach(pro=>{
          selectStr+=`${pro},`;
        });

        selectStr = selectStr.substr(0, selectStr.length-1);
      }
     
      return selectStr;
    }

    private static async HasPermssionOnOject(spHttpClient:SPHttpClient, apiUrl:string, permission:SP.PermissionKind)
    {
      var res = await spHttpClient.get(apiUrl, SPHttpClient.configurations.v1);
      var responseJson = await res.json();
      if(res.ok)
      {        
        console.log(`GetUserPermissions done for API url ${apiUrl}`);
        var permissions = new SP.BasePermissions();
        permissions.fromJson(responseJson);        
        var hasPermission = permissions.has(permission);      
        return hasPermission;
      }
      else
      {       
        console.log( `Failed GetUserPermissions for API url ${apiUrl}`);
        return false;
      }
    }
    
    public static async GetSiteChanges(spHttpClient:SPHttpClient,siteID:string,siteUrl:string,startDate:Date)
    {
      let apiUrl = `${siteUrl}/_api/site/getChanges`;

      // Set the ChangeTokenStart to two days ago to reduce how much data is returned from the change log. Depending on your requirements, you might want to change this value. 
      // The value of the string assigned to ChangeTokenStart.StringValue is semicolon delimited, and takes the following parameters in the order listed:
      // Version number. 
      // The change scope (0 - Content Database, 1 - site collection, 2 - site, 3 - list).
      // GUID of the item the scope applies to (for example, GUID of the list). 
      // Time (in UTC) from when changes occurred.
      // Initialize the change item on the ChangeToken using a default value of -1.

      
      //https://docs.microsoft.com/en-us/previous-versions/office/developer/sharepoint-2010/ee550385(v=office.14)
      //https://github.com/SharePoint/sp-dev-docs/issues/5964
      //https://docs.microsoft.com/en-us/sharepoint/dev/solution-guidance/query-sharepoint-change-log-with-changequery-and-changetoken
      //https://docs.microsoft.com/en-us/previous-versions/office/sharepoint-visio/jj245903(v%3doffice.15)

      //var startDate = new Date()
      //startDate.setDate(startDate.getDate() - 30);
      //let startDateUTCStr : number = startDate.getTime() + (startDate.getTimezoneOffset() * 60000);
      let body = {};
      if(startDate)
      {
        let startDateUTCStr : number = (startDate.getTime() * 10000) + 621355968000000000;
        var stringValue1 = `1;1;${siteID};${startDateUTCStr};-1`;
        body = { 
          'query': {
            //"Add":true,
            //"Update":true,
            //"Rename":true,
            "Item":true,
            "DeleteObject":true,
            'ChangeTokenStart':{
              'StringValue': `${stringValue1}`
            }
          }
        };
      }
      else{
        body = { 
          'query': {
            "Item":true,
            "DeleteObject":true
          }
        };
      }
    
      let spOpts = {  
            headers: {         
              "Content-Type": "application/json;odata=verbose"
            },  
            body: JSON.stringify(body)  
        };  
      var res = await spHttpClient.post(apiUrl, SPHttpClient.configurations.v1, spOpts);      
      if(res.ok)
      {
         var resJson = await res.json();
         console.log(`GetSiteChanges Done`);          
         return resJson;
      }
      else
      {
       console.log(`GetSiteChanges Failed - ${res.text()}`);       
      }
    }

    public static async GetListPath(spHttpClient:SPHttpClient, siteAbsoluteUrl:string, listID:string)
    {
      var apiUrl = `${siteAbsoluteUrl}/_api/web/lists/getbyId('${listID}')/RootFolder`;      
      var res = await spHttpClient.get(apiUrl, SPHttpClient.configurations.v1);
      if(res.ok)
      {
        var responseJson = await res.json();
        console.log(responseJson);
        console.log(`GetListPath done for API url ${apiUrl}`);
        return await responseJson.ServerRelativeUrl;
      }
      else
      {
        var message = `Failed GetListPath for API url ${apiUrl}`;
        console.log(message);
        Promise.reject(message);
        return null;
      }
    }

    public static async GetListbyId(spHttpClient:SPHttpClient, siteAbsoluteUrl:string, listID:string)
    {
      var apiUrl = `${siteAbsoluteUrl}/_api/web/lists/getbyId('${listID}')`;      
      var res = await spHttpClient.get(apiUrl, SPHttpClient.configurations.v1);
      if(res.ok)
      {
        var responseJson = await res.json();
        console.log(responseJson);
        console.log(`GetListPath done for API url ${apiUrl}`);
        return await responseJson;
      }
      else
      {
        var message = `Failed GetListPath for API url ${apiUrl}`;
        console.log(message);
        Promise.reject(message);
        return null;
      }
    }
    public static async GetDrives(spHttpClient:SPHttpClient, siteAbsoluteUrl:string)
    {
      var apiUrl = `${siteAbsoluteUrl}/_api/v2.0/drives?$select=id,name`;      
      var res = await spHttpClient.get(apiUrl, SPHttpClient.configurations.v1);
      if(res.ok)
      {
        var responseJson = await res.json();
        console.log(responseJson);
        console.log(`GetDrives done for API url ${apiUrl}`);
        return await responseJson.value;
      }
      else
      {
        var message = `Failed GetDrives for API url ${apiUrl}`;
        console.log(message);
        Promise.reject(message);
        return null;
      }
    }
    
    public static async IsObjectExisting(spHttpClient:SPHttpClient, apiUrl:string)
    {
        var objRes = await spHttpClient.get(apiUrl, SPHttpClient.configurations.v1);
        if(objRes.ok)
        {
            var resJson = await objRes.json();
            if(resJson.error)
            {
                console.error(`${resJson.error.message} ${apiUrl}.`);
                return false;
            }
            return true;
        }
        
        console.error(`Failed to get data from request ${apiUrl}.`);
        return false;    
    }    
}