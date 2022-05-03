import {SPHttpClient,ISPHttpClientOptions} from '@microsoft/sp-http';
import SPOQAHelper from './SPOQAHelper';

export default class FormsHelper
{
    public static async FixMissDisForm(spHttpClient:SPHttpClient, siteAbsoluteUrl:string, listTitle:string)
    {      
      var apiUrl = `${siteAbsoluteUrl}/_api/web/Lists/getByTitle('${listTitle}')`;
      var res = await spHttpClient.get(apiUrl, SPHttpClient.configurations.v1);
      var resJson = await res.json();
      console.log(resJson);       
      var listId = resJson.Id; // get list ID resJson.Id
      var baseTempl = resJson.BaseTemplate; //get list or library
      
      // get list root folder https://chengc.sharepoint.com/sites/abc/_api/web/Lists/getByTitle('LargeList')/rootFolder/      
      var rootfolderApi = apiUrl+"/rootFolder";
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
      
      var addFileApiUrl = `${siteAbsoluteUrl}/_api/web/GetFolderByServerRelativeUrl('${formPath}/Forms')/Files/Add(url='DispForm.aspx', overwrite=true)`;
      if(baseTempl != 101)
      {
        addFileApiUrl = `${siteAbsoluteUrl}/_api/web/GetFolderByServerRelativeUrl('${formPath}')/Files/Add(url='DispForm.aspx', overwrite=true)`;
      }

      let spOpts : ISPHttpClientOptions  = {
        headers: {
          "Accept": "application/json",
          "Content-Type": "application/json"
        },
        body: displayFormHtml        
      };
      
      var addFileRes = await spHttpClient.post(addFileApiUrl, SPHttpClient.configurations.v1, spOpts);
      return addFileRes;
    }

    
    public static async FixMissNewForm(spHttpClient:SPHttpClient, siteAbsoluteUrl:string, listTitle:string)
    {      
      var apiUrl = `${siteAbsoluteUrl}/_api/web/Lists/getByTitle('${listTitle}')`;
      var res = await spHttpClient.get(apiUrl, SPHttpClient.configurations.v1);
      var resJson = await res.json();
      console.log(resJson);       
      var listId = resJson.Id; // get list ID resJson.Id
      var baseTempl = resJson.BaseTemplate; //get list or library
      
      // get list root folder https://chengc.sharepoint.com/sites/abc/_api/web/Lists/getByTitle('LargeList')/rootFolder/      
      var rootfolderApi = apiUrl+"/rootFolder";
      var rootFolderRes = await spHttpClient.get(rootfolderApi, SPHttpClient.configurations.v1);
      var rootFolderResJson = await rootFolderRes.json();
      var formPath =  `${rootFolderResJson.ServerRelativeUrl}`;      
      
      var webpartId = SPOQAHelper.GenerateUUID();     
      var newFormHtml = `<%@ Page language="C#" MasterPageFile="~masterurl/default.master"    Inherits="Microsoft.SharePoint.WebPartPages.WebPartPage,Microsoft.SharePoint,Version=16.0.0.0,Culture=neutral,PublicKeyToken=71e9bce111e9429c" meta:webpartpageexpansion="full" meta:progid="SharePoint.WebPartPage.Document"  %>
        <%@ Register Tagprefix="SharePoint" Namespace="Microsoft.SharePoint.WebControls" Assembly="Microsoft.SharePoint, Version=16.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %> <%@ Register Tagprefix="Utilities" Namespace="Microsoft.SharePoint.Utilities" Assembly="Microsoft.SharePoint, Version=16.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %> <%@ Import Namespace="Microsoft.SharePoint" %> <%@ Assembly Name="Microsoft.Web.CommandUI, Version=16.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %> <%@ Register Tagprefix="WebPartPages" Namespace="Microsoft.SharePoint.WebPartPages" Assembly="Microsoft.SharePoint, Version=16.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
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
        <asp:Content ContentPlaceHolderId="PlaceHolderTitleAreaSeparator" runat="server"/>
        <asp:Content ContentPlaceHolderId="PlaceHolderPageImage" runat="server">
          <img src="/_layouts/15/images/blank.gif?rev=47" width='1' height='1' alt="" data-accessibility-nocheck="true"/>
        </asp:Content>
        <asp:Content ContentPlaceHolderId="PlaceHolderUtilityContent" runat="server">
        <SharePoint:ScriptBlock runat="server">var fCtl=false;
        function EnsureUploadCtl()
        {
          return browseris.ie5up && !browseris.mac &&
            null != document.getElementById("idUploadCtl");
        }
        function MultipleUploadView()
        {
          if (EnsureUploadCtl())
          {
            treeColor = GetTreeColor();
            document.all.idUploadCtl.SetTreeViewColor(treeColor);
            if(!fCtl)
            {
              rowsArr = document.all.formTbl.rows;
              for(i=0; i < rowsArr.length; i++)
              {
                if ((rowsArr[i].id != "OverwriteField") &&
                  (rowsArr[i].id != "trUploadCtl"))
                {
                  rowsArr[i].removeNode(true);
                  i=i-1;
                }
              }
              document.all.reqdFldTxt.removeNode(true);
              newCell = document.all.OverwriteField.insertCell();
              newCell.innerHTML = "&#160;";
              newCell.style.width="60%";
              document.all("dividMultipleView").style.display="inline";
              fCtl = true;
            }
          }
        }
        function RemoveMultipleUploadItems()
        {
          if(browseris.nav || browseris.mac ||
            !EnsureUploadCtl()
          )
          {
            formTblObj = document.getElementById("formTbl");
            if(formTblObj)
            {
              rowsArr = formTblObj.rows;
              for(i=0; i < rowsArr.length; i++)
              {
                if (rowsArr[i].id == "trUploadCtl" || rowsArr[i].id == "diidIOUploadMultipleLink")
                {
                  formTblObj.deleteRow(i);
                }
              }
            }
          }
        }
        function DocumentUpload()
        {
          if (fCtl)
          {
            document.all.idUploadCtl.MultipleUpload();
          }
          else
          {
            ClickOnce();
          }
        }
        function GetTreeColor()
        {
          var bkColor="";
          if(null != document.all("onetidNavBar"))
            bkColor = document.all.onetidNavBar.currentStyle.backgroundColor;
          if(bkColor=="")
          {
            numStyleSheets = document.styleSheets.length;
            for(i=numStyleSheets-1; i>=0; i--)
            {
              numRules = document.styleSheets(i).rules.length;
              for(ruleIndex=numRules-1; ruleIndex>=0; ruleIndex--)
              {
                if(document.styleSheets[i].rules.item(ruleIndex).selectorText==".ms-uploadcontrol")
                  uploadRule = document.styleSheets[i].rules.item(ruleIndex);
              }
            }
            if(uploadRule)
              bkColor = uploadRule.style.backgroundColor;
          }
          return(bkColor);
        } </SharePoint:ScriptBlock>
        <SharePoint:ScriptBlock runat="server">function _spBodyOnLoad()
          {
            var frm = document.forms[MSOWebPartPageFormName];
            frm.encoding="multipart/form-data";
          } </SharePoint:ScriptBlock>
        </asp:Content>
        <asp:Content ContentPlaceHolderId="PlaceHolderMain" runat="server">
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
          <ID>g_34b36896_da58_4683_abfe_a62cb92af8f2</ID>
          <ListName xmlns="http://schemas.microsoft.com/WebPart/v2/ListForm">{${listId}}</ListName>
          <ListId xmlns="http://schemas.microsoft.com/WebPart/v2/ListForm">${listId}</ListId>
          <PageType xmlns="http://schemas.microsoft.com/WebPart/v2/ListForm">PAGE_NEWFORM</PageType>
          <FormType xmlns="http://schemas.microsoft.com/WebPart/v2/ListForm">8</FormType>
          <ControlMode xmlns="http://schemas.microsoft.com/WebPart/v2/ListForm">New</ControlMode>
          <ViewFlag xmlns="http://schemas.microsoft.com/WebPart/v2/ListForm">1048576</ViewFlag>
          <ViewFlags xmlns="http://schemas.microsoft.com/WebPart/v2/ListForm">Default</ViewFlags>
          <ListItemId xmlns="http://schemas.microsoft.com/WebPart/v2/ListForm">0</ListItemId>
        </WebPart>
        </WebPartPages:ListFormWebPart>
        </ZoneTemplate></WebPartPages:WebPartZone>
          <input type="hidden" name="VTI-GROUP" value="0"/>
        </asp:Content>
        <asp:Content ContentPlaceHolderId="PlaceHolderPageTitle" runat="server">
          <SharePoint:EncodedLiteral runat="server" text="<%$Resources:wss,upload_pagetitle_form%>" EncodeMethod='HtmlEncode'/>
        </asp:Content>`;
      
      var addFileApiUrl = `${siteAbsoluteUrl}/_api/web/GetFolderByServerRelativeUrl('${formPath}/Forms')/Files/Add(url='Upload.aspx', overwrite=true)`;
      if(baseTempl != 101) // Not a document library
      {
        newFormHtml = `<%@ Page language="C#" MasterPageFile="~masterurl/default.master"    Inherits="Microsoft.SharePoint.WebPartPages.WebPartPage,Microsoft.SharePoint,Version=16.0.0.0,Culture=neutral,PublicKeyToken=71e9bce111e9429c" meta:webpartpageexpansion="full" meta:progid="SharePoint.WebPartPage.Document"  %>
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
          <ID>g_8678e685_f0c8_4c19_8572_003b3e0f1162</ID>
          <ListName xmlns="http://schemas.microsoft.com/WebPart/v2/ListForm">{${listId}}</ListName>
          <ListId xmlns="http://schemas.microsoft.com/WebPart/v2/ListForm">${listId}</ListId>
          <PageType xmlns="http://schemas.microsoft.com/WebPart/v2/ListForm">PAGE_NEWFORM</PageType>
          <FormType xmlns="http://schemas.microsoft.com/WebPart/v2/ListForm">8</FormType>
          <ControlMode xmlns="http://schemas.microsoft.com/WebPart/v2/ListForm">New</ControlMode>
          <ViewFlag xmlns="http://schemas.microsoft.com/WebPart/v2/ListForm">1048576</ViewFlag>
          <ViewFlags xmlns="http://schemas.microsoft.com/WebPart/v2/ListForm">Default</ViewFlags>
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
        <SharePoint:StyleBlock runat="server">
        .ms-bodyareaframe {
          padding: 8px;
          border: none;
        }
        </SharePoint:StyleBlock>
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
        addFileApiUrl = `${siteAbsoluteUrl}/_api/web/GetFolderByServerRelativeUrl('${formPath}')/Files/Add(url='NewForm.aspx', overwrite=true)`;
      }

      let spOpts : ISPHttpClientOptions  = {
        headers: {
          "Accept": "application/json",
          "Content-Type": "application/json"
        },
        body: newFormHtml        
      };
      
      var addFileRes = await spHttpClient.post(addFileApiUrl, SPHttpClient.configurations.v1, spOpts);
      return addFileRes;
    }

    public static async FixMissEditForm(spHttpClient:SPHttpClient, siteAbsoluteUrl:string, listTitle:string)
    {      
      var apiUrl = `${siteAbsoluteUrl}/_api/web/Lists/getByTitle('${listTitle}')`;
      var res = await spHttpClient.get(apiUrl, SPHttpClient.configurations.v1);
      var resJson = await res.json();
      console.log(resJson);       
      var listId = resJson.Id; // get list ID resJson.Id
      var baseTempl = resJson.BaseTemplate; //get list or library
      
      // get list root folder https://chengc.sharepoint.com/sites/abc/_api/web/Lists/getByTitle('LargeList')/rootFolder/      
      var rootfolderApi = apiUrl+"/rootFolder";
      var rootFolderRes = await spHttpClient.get(rootfolderApi, SPHttpClient.configurations.v1);
      var rootFolderResJson = await rootFolderRes.json();
      var formPath =  `${rootFolderResJson.ServerRelativeUrl}`;      
      
      var webpartId = SPOQAHelper.GenerateUUID();
      var editFormHtml = `<%@ Page language="C#" MasterPageFile="~masterurl/default.master"    Inherits="Microsoft.SharePoint.WebPartPages.WebPartPage,Microsoft.SharePoint,Version=16.0.0.0,Culture=neutral,PublicKeyToken=71e9bce111e9429c" meta:webpartpageexpansion="full" meta:progid="SharePoint.WebPartPage.Document"  %>
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
        <ID>g_d264b632_9aa8_4c68_be91_1cec21099db8</ID>
        <ListName xmlns="http://schemas.microsoft.com/WebPart/v2/ListForm">{${listId}}</ListName>
        <ListId xmlns="http://schemas.microsoft.com/WebPart/v2/ListForm">${listId}</ListId>
        <PageType xmlns="http://schemas.microsoft.com/WebPart/v2/ListForm">PAGE_EDITFORM</PageType>
        <FormType xmlns="http://schemas.microsoft.com/WebPart/v2/ListForm">6</FormType>
        <ControlMode xmlns="http://schemas.microsoft.com/WebPart/v2/ListForm">Edit</ControlMode>
        <ViewFlag xmlns="http://schemas.microsoft.com/WebPart/v2/ListForm">1048576</ViewFlag>
        <ViewFlags xmlns="http://schemas.microsoft.com/WebPart/v2/ListForm">Default</ViewFlags>
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
      <SharePoint:StyleBlock runat="server">
      .ms-bodyareaframe {
        padding: 8px;
        border: none;
      }
      </SharePoint:StyleBlock>
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
      
      var addFileApiUrl = `${siteAbsoluteUrl}/_api/web/GetFolderByServerRelativeUrl('${formPath}/Forms')/Files/Add(url='EditForm.aspx', overwrite=true)`;
      if(baseTempl != 101)
      {
        addFileApiUrl = `${siteAbsoluteUrl}/_api/web/GetFolderByServerRelativeUrl('${formPath}')/Files/Add(url='EditForm.aspx', overwrite=true)`;
      }
      
      let spOpts : ISPHttpClientOptions  = {
        headers: {
          "Accept": "application/json",
          "Content-Type": "application/json"
        },
        body: editFormHtml        
      };
      
      var addFileRes = await spHttpClient.post(addFileApiUrl, SPHttpClient.configurations.v1, spOpts);
      return addFileRes;
    }
}