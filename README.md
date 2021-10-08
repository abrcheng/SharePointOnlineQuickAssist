# SharePointOnlineQuickAssist-Tutorial Materials

SharePoint Online Quick Assist is a SPFX webpart that appears inside a SharePoint page in the browser. Site administrators could use the tool to diagnose some common issues and fix them.




## Available features

* Search Issue For A Specific Document


  This feature helps user diagnose the issue when a specific document does not appear in the search results

  <IMG src=.\assets\SearchSpecificDocument.JPG>

   [More details](https://github.com/abrcheng/SharePointOnlineQuickAssist/blob/main/SPFX/SPOQA/SearchSpecificDocument.md)







* Job Title Sync Issue
	

  This feature helps validate user's 'job title' in AAD, SPO user profile and site.
  
  <IMG src=.\assets\JobTitleSync.JPG>
  
   [More details](https://github.com/abrcheng/SharePointOnlineQuickAssist/blob/main/SPFX/SPOQA/JobTitleSyncIssue.md)
    
*  Photo Sync Issue
   This feature helps user compare their profile photo from AAD to SPO user profile.
    
  <IMG src=.\assets\CheckUserProfilePhoto.JPG>
    
    
## Using the tool

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
* gulp clean
* gulp bundle --ship
* gulp package-solution --ship
* Upload .sppkg file from sharepoint\solution to your tenant App Catalog
	* E.g.: https://&lt;tenant&gt;.sharepoint.com/sites/AppCatalog/AppCatalog
<IMG src=.\assets\UploadSolution.JPG>
	  
* Approve API access requests in SharePoint admin center  
        * https://&lt;tenant&gt;-admin.sharepoint.com/_layouts/15/online/AdminHome.aspx#/webApiPermissionManagement 
<IMG src=.\assets\ApproveAPI.JPG>	
	
* Add the web part to a site collection, and test it on a page    
    
### Additional resources

* [Overview of the SharePoint Framework](https://docs.microsoft.com/sharepoint/dev/spfx/sharepoint-framework-overview)
* [SharePoint Framework development tools and libraries](https://docs.microsoft.com/sharepoint/dev/spfx/tools-and-libraries)
* [Getting Started](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)
