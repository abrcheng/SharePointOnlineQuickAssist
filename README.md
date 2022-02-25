# SharePointOnlineQuickAssist-Tutorial Materials

SharePoint Online Quick Assist is a SPFX webpart that appears inside a SharePoint page in the browser. Site administrators could use the tool to diagnose some common issues and fix them.

This tool is provided by the copyright holders and contributors “as is” and any express or implied warranties, including, but not limited to, the implied warranties of merchantability and fitness for a particular purpose are disclaimed. In no event shall the copyright owner or contributors be liable for any direct, indirect, incidental, special, exemplary, or consequential damages (including, but not limited to, procurement of substitute goods or services; loss of use, data, or profits; or business interruption) however caused and on any theory of liability, whether in contract, strict liability, or tort (including negligence or otherwise) arising in any way out of the use of this software, even if advised of the possibility of such damage.

## Please note, if you want to use the auto fix function on the affected site then the custom script for the site should be enabled https://docs.microsoft.com/en-us/sharepoint/allow-or-prevent-custom-script#to-allow-custom-script-on-other-sharepoint-sites

## Available features

* **Search Issue For A Specific Document**


  This feature helps user diagnose the issue when a specific document does not appear in the search results

  <IMG src=.\assets\NoCrawl.JPG>

   [More details](https://github.com/abrcheng/SharePointOnlineQuickAssist/blob/main/SPFX/SPOQA/SearchSpecificDocument.md)


* **Search Issue For A Specific Site**
	
  This feature helps user diagnose the issue when a specific site does not appear in the search results	

  <IMG src=.\assets\SiteNoCrawl.JPG>
	  
   [More details](https://github.com/abrcheng/SharePointOnlineQuickAssist/blob/main/SPFX/SPOQA/SearchSite.md)
	  

* **Job Title Sync Issue**
	

  This feature helps validate user's 'job title' in AAD, SPO user profile and site.
  
  <IMG src=.\assets\JobTitle.JPG>
  
   [More details](https://github.com/abrcheng/SharePointOnlineQuickAssist/blob/main/SPFX/SPOQA/JobTitleSyncIssue.md)
    
*  **Photo Sync Issue**
   This feature helps user compare their profile photo from AAD to SPO user profile.
    
  <IMG src=.\assets\CheckUserProfilePhoto.JPG>
    
* **User Info list sync issue**,
	User/Group mail haven’t been synced to user info list caused the mail can’t be send to User/Group in user alert, workflow 
	User/Group display name updated but haven’t synced to user info list cause mismatch issue 
	User’ phone number haven’t been synced to user info list cause mismatch issue 
	User’ job title haven’t been synced to user info list cause mismatch issue 
![image](https://user-images.githubusercontent.com/21354416/144986960-e4befdd6-b9d6-40a0-bb54-fc90ca8d0d70.png)
![image](https://user-images.githubusercontent.com/21354416/144987002-085d0652-f243-4c29-84b6-94452d3afdee.png)

*  **OneDrive library sync issue**,
	 OneDrive sync button can't be found.
	 Library synced as read only 
	 ![image](https://user-images.githubusercontent.com/21354416/144987185-d18e2e24-5b35-4436-ba3f-002dd95819c7.png)
         ![image](https://user-images.githubusercontent.com/21354416/144987204-83e0eb28-9f10-4ec4-858a-cc31c78d0d32.png)
	  
*  **Missing New/Display/Edit forms issue**
	  ![image](https://user-images.githubusercontent.com/21354416/147523270-bef520ae-e487-4414-a449-6348b31efe82.png)
[More details](https://github.com/abrcheng/SharePointOnlineQuickAssist/releases/tag/1.21.12.28)
	  
*  **Filter and restore items from recycle bin**
	  ![image](https://user-images.githubusercontent.com/21354416/155690547-1cb262ef-30d7-4717-9fa0-5cd5c6a4d8dd.png)	  
[More details](https://github.com/abrcheng/SharePointOnlineQuickAssist/releases/tag/1.22.02.25)

	  
## Install the tool to SharePoint Online 
* Upload SPOQA.sppkg from https://github.com/abrcheng/SharePointOnlineQuickAssist/blob/main/Packages/spoqa.sppkg to your tenant App Catalog
	* E.g.: https://&lt;tenant&gt;.sharepoint.com/sites/AppCatalog/AppCatalog
<IMG src=.\assets\UploadSolution.JPG>

* Deploy the app when you see the prompt as follow
<IMG src=.\assets\Deploy.JPG>	
	
	  
* Approve API access requests in SharePoint admin center  
        * https://&lt;tenant&gt;-admin.sharepoint.com/_layouts/15/online/AdminHome.aspx#/webApiPermissionManagement 
<IMG src=.\assets\ApproveAPI.JPG>	
	
* Add the web part to a site collection, and test it on a page    
<IMG src=.\assets\WebPart.JPG>	
    
## If you want to contribute/customzied this tool, you may try below steps,
	
To build and start using these projects, you'll need to clone and build the projects.

Clone this repository by executing the following command in your console:

```shell
git clone https://github.com/abrcheng/SharePointOnlineQuickAssist.git
```

Navigate to the cloned repository folder which should be the same as the repository name:

```shell
cd SharePointOnlineQuickAssist
```

To access the webpart use the following command.

```shell
cd SPFX
cd SPOQA
```


Now run the following command to install the npm packages:

```shell
npm install
```

This will install the required npm packages and dependencies to build and run the client-side project.

Once the npm packages are installed, run the following command to preview your web parts in SharePoint Workbench:

```shell
gulp serve
```

### Deploy
```shell	  
gulp clean
gulp bundle --ship
gulp package-solution --ship
```
### Additional resources

* [Overview of the SharePoint Framework](https://docs.microsoft.com/sharepoint/dev/spfx/sharepoint-framework-overview)
* [SharePoint Framework development tools and libraries](https://docs.microsoft.com/sharepoint/dev/spfx/tools-and-libraries)
* [Getting Started](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)
