import { WebPartContext } from "@microsoft/sp-webpart-base"; 
import {MSGraphClient, SPHttpClient} from '@microsoft/sp-http';
import {SPUser} from '@microsoft/sp-page-context';

export interface IExoQuickAssistProps {
  msGraphClient:MSGraphClient;
  // spHttpClient:SPHttpClient;
  // webAbsoluteUrl:string;
  // webUrl:string;
  // rootUrl:string;
  currentUser:SPUser;
  ctx:WebPartContext;
}
