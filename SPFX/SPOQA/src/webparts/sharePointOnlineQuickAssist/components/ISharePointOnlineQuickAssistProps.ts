import { WebPartContext } from "@microsoft/sp-webpart-base"; 
import {MSGraphClient, SPHttpClient} from '@microsoft/sp-http';
export interface ISharePointOnlineQuickAssistProps {  
  msGraphClient:MSGraphClient;
  spHttpClient:SPHttpClient;
  webAbsoluteUrl:string;
  webUrl:string;
  rootUrl:string;
}
