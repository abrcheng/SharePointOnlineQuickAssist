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
        var message = "Failed to get data from graph API";
        console.log(message);
        Promise.reject(message);
      }
    }
}