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
      var apiUrl = `${webAbsoluteUrl}/_api/web/lists/getByTitle('User Information List')/items?$filter=Name eq '${encodeURIComponent(account)}'`;      
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
    public static async FixJobTitleInUserInfoList(userId:number, spHttpClient:SPHttpClient, webAbsoluteUrl:string, newJobTitle:string, successCallBack:Function, failedCallback:Function) {
       /*var userItem = {
        "__metadata": {"type":"SP.Data.UserInfoItem"},        
        "Title":"SPOQA",
         "JobTitle": newJobTitle,        
         "Department":"UpdateBySPOQA"
       };
       
       var userItemBody =  JSON.stringify(userItem);
       // const apiUrl = `${webAbsoluteUrl}/_api/Web/SiteUserInfoList/Items(${userId})`;
       const apiUrl = `${webAbsoluteUrl}/_api/web/siteusers/getbyid(${userId})`;
       const options: ISPHttpClientOptions = {     
        headers: {
          "Accept": "application/json;odata=verbose",
          "Content-Type": "application/json;odata=verbose",
          "OData-Version": "" ,//Really important to specify,
          "X-HTTP-Method": 'MERGE',
          'IF-MATCH': '*'
        },
        body: userItemBody
      };

      var res = await spHttpClient.post(apiUrl, SPHttpClient.configurations.v1, options);
      if(res.ok)
      {
        var responseJson = await res.json();
        console.log(`FixJobTitleInUserInfoList done for API url ${apiUrl}`);          
        return await responseJson;
      }
      else
      {
        var message = `Failed FixJobTitleInUserInfoList for API url ${apiUrl}`;
        console.log(message);
        Promise.reject(message);
      }*/    

      const context: SP.ClientContext = new SP.ClientContext(webAbsoluteUrl);     
      const userItem: SP.ListItem = context.get_web().get_lists().getByTitle("User Information List").getItemById(userId); 
      userItem.set_item('JobTitle', newJobTitle);      
      userItem.update();
      context.executeQueryAsync((sender: any, args: SP.ClientRequestSucceededEventArgs): void => {successCallBack();}, (sender: any, args: SP.ClientRequestSucceededEventArgs): void => {failedCallback();});      
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
    
    // https://chengc.sharepoint.com/_api/web/lists/getByTitle('User Information List')/items?$filter=Name eq 'i:0#.f|membership|abc@chengc.onmicrosoft.com'
    
}