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

    public static async FixMissDisForm(spHttpClient:SPHttpClient, siteAbsoluteUrl:string, listTitle:string)
    {      
      var apiUrl = `${siteAbsoluteUrl}/_api/web/Lists/getByTitle('${listTitle}')`;
      var res = await spHttpClient.get(apiUrl, SPHttpClient.configurations.v1);
      var resJson = await res.json();
      console.log(resJson);       
      var listId = resJson.Id; // get list ID resJson.Id
      
      // get list root folder https://chengc.sharepoint.com/sites/abc/_api/web/Lists/getByTitle('LargeList')/rootFolder/      
      var rootfolderApi = apiUrl+"/rootFolder"
      var rootFolderRes = await spHttpClient.get(rootfolderApi, SPHttpClient.configurations.v1);
      var rootFolderResJson = await rootFolderRes.json();
      var formPath =  `${rootFolderResJson.ServerRelativeUrl}`;      
      
      var webpartId = SPOQAHelper.GenerateUUID();
      var displayFormHtml = `<%@ Page language="C#" MasterPageFile="~masterurl/default.master"    Inherits="Microsoft.SharePoint.WebPartPages.WebPartPage,Microsoft.SharePoint,Version=16.0.0.0,Culture=neutral,PublicKeyToken=71e9bce111e9429c" meta:webpartpageexpansion="full" meta:progid="SharePoint.WebPartPage.Document"  %>
      <%@ Register Tagprefix="SharePoint" Namespace="Microsoft.SharePoint.WebControls" Assembly="Microsoft.SharePoint, Version=16.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %> <%@ Register Tagprefix="Utilities" Namespace="Microsoft.SharePoint.Utilities" Assembly="Microsoft.SharePoint, Version=16.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %> <%@ Import Namespace="Microsoft.SharePoint" %> <%@ Assembly Name="Microsoft.Web.CommandUI, Version=16.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %> <%@ Register Tagprefix="WebPartPages" Namespace="Microsoft.SharePoint.WebPartPages" Assembly="Microsoft.SharePoint, Version=16.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
      <asp:Content ContentPlaceHolderId="PlaceHolderPageTitle" runat="server">
        <SharePoint:ListFormPageTitle runat="server"/>
      </asp:Content>
      <asp:Content ContentPlaceHolderId="PlaceHolderPageTitleInTitleArea" runat="server">
        <span class="die">
          <SharePoint:ListProperty Property="LinkTitle" runat="server" id="ID_LinkTitle"/>
        </span>
      </asp:Content>
      <asp:Content ContentPlaceHolderId="PlaceHolderPageImage" runat="server">
        <img src="/_layouts/15/images/blank.gif?rev=47" width='1' height='1' alt="" data-accessibility-nocheck="true"/>
      </asp:Content>
      <asp:Content ContentPlaceHolderId="PlaceHolderMain" runat="server">
      <SharePoint:UIVersionedContent UIVersion="4" runat="server">
        <ContentTemplate>
        <div style="padding-left:5px">
        </ContentTemplate>
      </SharePoint:UIVersionedContent>
        <table class="ms-core-tableNoSpace" id="onetIDListForm" role="presentation">
         <tr>
          <td>
         <WebPartPages:WebPartZone runat="server" FrameType="None" ID="Main" Title="loc:Main"><ZoneTemplate>
      <WebPartPages:ListFormWebPart runat="server" __MarkupType="xmlmarkup" WebPart="true" __WebPartId="{${webpartId}}" >
      <WebPart xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns="http://schemas.microsoft.com/WebPart/v2">
        <Title>${listTitle}</Title>
        <FrameType>Default</FrameType>
        <Description />
        <IsIncluded>true</IsIncluded>
        <PartOrder>2</PartOrder>
        <FrameState>Normal</FrameState>
        <Height />
        <Width />
        <AllowRemove>true</AllowRemove>
        <AllowZoneChange>true</AllowZoneChange>
        <AllowMinimize>true</AllowMinimize>
        <AllowConnect>true</AllowConnect>
        <AllowEdit>true</AllowEdit>
        <AllowHide>true</AllowHide>
        <IsVisible>true</IsVisible>
        <DetailLink />
        <HelpLink />
        <HelpMode>Modeless</HelpMode>
        <Dir>Default</Dir>
        <PartImageSmall />
        <MissingAssembly>Cannot import this Web Part.</MissingAssembly>
        <PartImageLarge />
        <IsIncludedFilter />
        <ExportControlledProperties>true</ExportControlledProperties>
        <ConnectionID>00000000-0000-0000-0000-000000000000</ConnectionID>
        <ID>g_782eb163_0b9e_4d6c_ba14_89a17fba9c75</ID>
        <ListName xmlns="http://schemas.microsoft.com/WebPart/v2/ListForm">{${listId}}</ListName>
        <ListId xmlns="http://schemas.microsoft.com/WebPart/v2/ListForm">${listId}</ListId>
        <PageType xmlns="http://schemas.microsoft.com/WebPart/v2/ListForm">PAGE_DISPLAYFORM</PageType>
        <FormType xmlns="http://schemas.microsoft.com/WebPart/v2/ListForm">4</FormType>
        <ControlMode xmlns="http://schemas.microsoft.com/WebPart/v2/ListForm">Display</ControlMode>
        <ViewFlag xmlns="http://schemas.microsoft.com/WebPart/v2/ListForm">0</ViewFlag>
        <ListItemId xmlns="http://schemas.microsoft.com/WebPart/v2/ListForm">0</ListItemId>
      </WebPart>
      </WebPartPages:ListFormWebPart>
          </ZoneTemplate></WebPartPages:WebPartZone>
          </td>
         </tr>
        </table>
      <SharePoint:UIVersionedContent UIVersion="4" runat="server">
        <ContentTemplate>
        </div>
        </ContentTemplate>
      </SharePoint:UIVersionedContent>
      </asp:Content>
      <asp:Content ContentPlaceHolderId="PlaceHolderAdditionalPageHead" runat="server">
        <SharePoint:DelegateControl runat="server" ControlId="FormCustomRedirectControl" AllowMultipleControls="true"/>
        <SharePoint:UIVersionedContent UIVersion="4" runat="server"><ContentTemplate>
          <SharePoint:CssRegistration Name="forms.css" runat="server"/>
        </ContentTemplate></SharePoint:UIVersionedContent>
      </asp:Content>
      <asp:Content ContentPlaceHolderId="PlaceHolderTitleLeftBorder" runat="server">
      <table cellpadding="0" height="100%" width="100%" cellspacing="0">
       <tr><td class="ms-areaseparatorleft"><img src="/_layouts/15/images/blank.gif?rev=47" width='1' height='1' alt="" data-accessibility-nocheck="true"/></td></tr>
      </table>
      </asp:Content>
      <asp:Content ContentPlaceHolderId="PlaceHolderTitleAreaClass" runat="server">
      <script type="text/javascript" id="onetidPageTitleAreaFrameScript">
        if (document.getElementById("onetidPageTitleAreaFrame") != null)
        {
          document.getElementById("onetidPageTitleAreaFrame").className="ms-areaseparator";
        }
      </script>
      </asp:Content>
      <asp:Content ContentPlaceHolderId="PlaceHolderBodyAreaClass" runat="server">
      <SharePoint:StyleBlock runat="server">.ms-bodyareaframe {
        padding: 8px;
        border: none;
      } </SharePoint:StyleBlock>
      </asp:Content>
      <asp:Content ContentPlaceHolderId="PlaceHolderBodyLeftBorder" runat="server">
      <div class='ms-areaseparatorleft'><img src="/_layouts/15/images/blank.gif?rev=47" width='8' height='100%' alt="" data-accessibility-nocheck="true"/></div>
      </asp:Content>
      <asp:Content ContentPlaceHolderId="PlaceHolderTitleRightMargin" runat="server">
      <div class='ms-areaseparatorright'><img src="/_layouts/15/images/blank.gif?rev=47" width='8' height='100%' alt="" data-accessibility-nocheck="true"/></div>
      </asp:Content>
      <asp:Content ContentPlaceHolderId="PlaceHolderBodyRightMargin" runat="server">
      <div class='ms-areaseparatorright'><img src="/_layouts/15/images/blank.gif?rev=47" width='8' height='100%' alt="" data-accessibility-nocheck="true"/></div>
      </asp:Content>
      <asp:Content ContentPlaceHolderId="PlaceHolderTitleAreaSeparator" runat="server"/>`;
      
       var addFileApiUrl = `${siteAbsoluteUrl}/_api/web/GetFolderByServerRelativeUrl('${formPath}')/Files/Add(url='DispForm.aspx', overwrite=true)`;
       let spOpts : ISPHttpClientOptions  = {
        headers: {
          "Accept": "application/json",
          "Content-Type": "application/json"
        },
        body: displayFormHtml        
      };
      
      var addFileRes = await spHttpClient.post(addFileApiUrl, SPHttpClient.configurations.v1, spOpts)
      return addFileRes;
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
            let apiUrl:string = `${siteAbsoluteUrl}/_api/web/GetFileByUrl('${relativeDocPath}')/Publish('Published by SharePointOnlineQuickAssist')`;           
            var res = await spHttpClient.post(apiUrl, SPHttpClient.configurations.v1, spOpts);   
            console.log(`${apiUrl} OK?${res.ok}`);
          }
     

      return resJson;
    }
}