import {SPHttpClient,ISPHttpClientOptions} from '@microsoft/sp-http';

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

    public static async GetSiteSearchResult(spHttpClient:SPHttpClient, webAbsoluteUrl:string)
    { 
      var apiUrl = `${webAbsoluteUrl}/_api/search/query?querytext=%27path:%22${webAbsoluteUrl}%22%20ContentClass:STS_Web%27`;
      var res = await RestAPIHelper.CallGetRest(apiUrl, spHttpClient);
      return res;
    }
}