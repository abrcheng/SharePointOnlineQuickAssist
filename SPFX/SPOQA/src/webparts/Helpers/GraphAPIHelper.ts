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
}