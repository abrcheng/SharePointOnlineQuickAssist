import { WebPartContext } from "@microsoft/sp-webpart-base";  
export default class ContextHelper
{
       public static context:any;
       public static GetInstance():WebPartContext
       {
           return ContextHelper.context;
       }
       public static SetInstace(webpartContext:WebPartContext)       
       {    // window.wpContext = webpartContext;            
            ContextHelper.context = webpartContext;
       }
}