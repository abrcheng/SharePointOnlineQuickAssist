# Depolyment Approaches  
## Depoly the tool from SharePoint Store 
1. Open the tenant app manage page https://**contoso**-admin.sharepoint.com/_layouts/15/tenantAppCatalog.aspx and click "SharePoint Store"

![image](https://user-images.githubusercontent.com/21354416/181446459-80487868-73a1-44dc-8732-aab11b36a1c9.png)

2. Search the "SharePoint Online Quick Assist"

![image](https://user-images.githubusercontent.com/21354416/181446926-71f61529-3b33-4c0b-8a58-a3d8328520c5.png)

3. Click the searched "SharePoint Online Quick Assist" and click "Add to app catalog" button

![image](https://user-images.githubusercontent.com/21354416/181447089-484ed56a-b2f1-4e95-8c76-7ad97c613492.png)

4. Click "Add" button for the "Confirm data access"

![image](https://user-images.githubusercontent.com/21354416/181447204-359d1818-1853-4b6a-897b-d4a695e20cb6.png)

5. Click "Go to API access page" button 

![image](https://user-images.githubusercontent.com/21354416/181447325-f52cd82c-ca38-4968-ba70-5045751478db.png)

6. Select pending API requests for "SharePoint Online Quick Assist" and click approve button

![image](https://user-images.githubusercontent.com/21354416/181447412-1c2ba036-e8fb-4030-ac15-06511b81239d.png)

7. Go to the site and add "SharePoint Online Quick Assist"

![image](https://user-images.githubusercontent.com/21354416/181447526-bf2d3ce3-e5f0-46cc-b548-d8833a01b6c3.png)

8. Add the "SharePoint Online Quick Assist" web part into page

![image](https://user-images.githubusercontent.com/21354416/181447638-5ab748de-865b-4b7f-a260-f775f7daa0b3.png)

## Deploy the tool tenant level app catalog 
* Upload SPOQA.sppkg from https://github.com/abrcheng/SharePointOnlineQuickAssist/blob/main/Packages/spoqa.sppkg to your tenant App Catalog
	* E.g.: https://&lt;tenant&gt;.sharepoint.com/sites/AppCatalog/AppCatalog
<IMG src=..\..\assets\UploadSolution.JPG>

* Deploy the app when you see the prompt as follow
<IMG src=..\..\assets\Deploy.JPG>	
	
	  
* Approve API access requests in SharePoint admin center  
        * https://&lt;tenant&gt;-admin.sharepoint.com/_layouts/15/online/AdminHome.aspx#/webApiPermissionManagement 
<IMG src=..\..\assets\ApproveAPI.JPG>	
	
* Add the web part to a site collection, and test it on a page    
<IMG src=..\..\assets\WebPart.JPG>	
	
## Deploy the tool site collection level app catalog 
Download and install SharePoint Online Management Shell.
* Open it and run the following: (You need Global admin or SharePoint admin rights. )
* Connect-SPOService https://contoso-admin.sharepoint.com
* Set-SPOSite -Identity https://contoso.sharepoint.com/sites/ASite -DenyAddAndCustomizePages 0
* Add-SPOSiteCollectionAppCatalog -Site https://contoso.sharepoint.com/sites/ASite
* Download the tool https://github.com/abrcheng/SharePointOnlineQuickAssist/blob/main/Packages/spoqa.sppkg. 
* Access https://contoso.sharepoint.com/sites/ASite/AppCatalog.
* Click “Upload” and upload “spoqa.sppkg”. 
* Click “Deploy” and click “Trust” button.
* Approve API access requests in SharePoint admin center,
        https://<tenant>-admin.sharepoint.com/_layouts/15/online/AdminHome.aspx#/webApiPermissionManagement
* Back to the site https://contoso.sharepoint.com/sites/ASite, click “Add an app”. 
* Add “spoqa-client-side-solution”.
* Add a new page in the site and add the web part “SharePoint Online Quick Assist”. “Publish” the page. 

        

   
