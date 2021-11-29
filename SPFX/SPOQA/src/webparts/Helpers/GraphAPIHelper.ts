import {MSGraphClient,SPHttpClient} from '@microsoft/sp-http';

export default class GraphAPIHelper
{
    public static async GetUserInfo(user: string, msGraphClient:MSGraphClient)
    {      
      var res = await msGraphClient.api(`/users/${user}`).get();
      if(res)
      {               
        console.log(`GraphAPIHelper.GetUserInfo for user ${user} done.`);
        return await res;
      }
      else
      {
        var message = `Failed to get uesr ${user} from graph API`;
        console.log(message);
        Promise.reject(message);
      }
    }

    public static async GetUserPhoto(user:string, msGraphClient:MSGraphClient)
    {
      var res = await msGraphClient.api(`/users/${user}/photo/$value`).responseType('blob').get();
      if(res)
      {               
        console.log(`GraphAPIHelper.GetUserPhoto for user ${user} done.`);
        return await res;
      }
      else
      {
        console.log("GraphAPIHelper.GetUserPhoto failed");
      }
    }
    
    public static async GetGroupMembers(groupid:string, msGraphClient:MSGraphClient)
    {      
      var res = await msGraphClient.api(`/groups/${groupid}/members?$select=id,mail`).get();
      if(res)
      {               
        console.log(`GraphAPIHelper.GetGroupMembers for user ${groupid} done.`);
        const graphResponse: any = res.value; 
        return await graphResponse;
      }
      else
      {
        var message = `Failed to get uesr ${groupid} from graph API`;
        console.log(message);
        Promise.reject(message);
      }
    }

    public static async AddUserinMembers(groupid:string, msGraphClient:MSGraphClient, useremail:string)
    {

      var resUserInfo = await msGraphClient.api(`/me`).get();
      console.log(`User info2: ${await resUserInfo["id"]}`);
      var userId = await resUserInfo["id"];

      //var keyOdataId = `@odata.id`;
      //var valueODataId = `https://graph.microsoft.com/v1.0/directoryObjects/${useremail}`;

      //const directoryObject = `@"{{ ""${keyOdataId}"": ""${valueODataId}"" }}`;

      var body: string = JSON.stringify({
        "@odata.id": `https://graph.microsoft.com/v1.0/directoryObjects/${userId}`
      });
      
      var res = await msGraphClient.api(`/groups/${groupid}/members/$ref`).post(body);
      if(res)
      {               
        console.log(`GraphAPIHelper.AddUserinMembers for user ${useremail} to ${groupid} done.`);
        const graphResponse: any = res.value; 
        return await graphResponse;
      }
      else
      {
        var message = `Failed to add uesr ${useremail} to ${groupid} via graph API`;
        console.log(message);
        Promise.reject(message);
      }
    }

    public static async GetGroupByEmail(msGraphClient:MSGraphClient, groupMail:string)
    {
      var res = await msGraphClient.api(`/groups?$filter=mail eq '${groupMail}'`).get();
      if(res)
      {               
        console.log(`GraphAPIHelper.GetGroupByEmail for group ${groupMail} done, get ${res.value.length} groups.`);
        return await res;
      }
      else
      {
        var message = `Failed to get group ${groupMail} from graph API`;
        console.log(message);
        Promise.reject(message);
      }
    }

    public static async GetUserByEmail(msGraphClient:MSGraphClient, userMail:string)
    {
      var res = await msGraphClient.api(`/users?$filter=mail eq '${userMail}'`).get();
      if(res)
      {               
        console.log(`GraphAPIHelper.GetUserByEmail for user ${userMail} done, get ${res.value.length} users`);
        return await res;
      }
      else
      {
        var message = `Failed to get uesr ${userMail} from graph API`;
        console.log(message);
        Promise.reject(message);
      }
    }
}