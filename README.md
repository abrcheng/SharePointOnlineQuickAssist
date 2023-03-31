# SharePointOnlineQuickAssist-Tutorial Materials

SharePoint Online Quick Assist is a SPFX webpart that appears inside a SharePoint page in the browser. Site administrators could use the tool to diagnose some common issues and fix them.

This tool is provided by the copyright holders and contributors “as is” and any express or implied warranties, including, but not limited to, the implied warranties of merchantability and fitness for a particular purpose are disclaimed. In no event shall the copyright owner or contributors be liable for any direct, indirect, incidental, special, exemplary, or consequential damages (including, but not limited to, procurement of substitute goods or services; loss of use, data, or profits; or business interruption) however caused and on any theory of liability, whether in contract, strict liability, or tort (including negligence or otherwise) arising in any way out of the use of this software, even if advised of the possibility of such damage.

## Please note, if you want to use the auto fix function on the affected site then the custom script for the site should be enabled https://docs.microsoft.com/en-us/sharepoint/allow-or-prevent-custom-script#to-allow-custom-script-on-other-sharepoint-sites

https://user-images.githubusercontent.com/89838160/169972855-a7605e68-c6b1-4235-af61-a4226bee5b45.mp4

# M365 First Aid x Hackathon 2022
[![M365 First Aid x Hackathon](https://user-images.githubusercontent.com/89838160/186566847-f6585954-413c-4b46-b39a-29087022f1ae.png)](https://www.youtube.com/watch?v=IBx8SICGFJE)

# M365 First Aid Introduction
[![IMAGE ALT TEXT HERE](https://user-images.githubusercontent.com/89838160/186567066-16e1338c-4236-40b8-b45d-d3be503ff9f9.PNG)](https://www.youtube.com/watch?v=lA31ikfL8JY)

## Available features

* **Restore Items**
This feature helps user to filter/restore/export items from SharePoint Online recycle bin,
![image](https://user-images.githubusercontent.com/21354416/155688589-199fc965-1333-4073-82b6-677444497a36.png)

[More details](https://github.com/abrcheng/SharePointOnlineQuickAssist/blob/main/Documents/RestoreItems.md)

* **Check Permssion issue**
This feature helps user to diagnose permssion issues (get access denied or 404 when accessing shared links even the linked document does existing),
![image](https://user-images.githubusercontent.com/21354416/160541128-832029ee-cfda-4f4a-913f-9512c43aaf67.png)

[More details](https://github.com/abrcheng/SharePointOnlineQuickAssist/blob/main/Documents/CheckPermissionIssue.md)

* **Search Issue For A Specific Document**


  This feature helps user to diagnose the issue when a specific document does not appear in the search results

  <IMG src=.\assets\NoCrawl.JPG>

   [More details](https://github.com/abrcheng/SharePointOnlineQuickAssist/blob/main/Documents/SearchSpecificDocument.md)


* **Search Issue For A Specific Site**
	
  This feature helps user diagnose the issue when a specific site does not appear in the search results	

  <IMG src=.\assets\SiteNoCrawl.JPG>
	  
   [More details](https://github.com/abrcheng/SharePointOnlineQuickAssist/blob/main/Documents/SearchSpecificSite.md)
	  

* **Job Title Sync Issue**
	

  This feature helps validate user's 'job title' in AAD, SPO user profile and site.
  
  <IMG src=.\assets\JobTitle.JPG>
  
   [More details](https://github.com/abrcheng/SharePointOnlineQuickAssist/blob/main/Documents/JobTitleSyncIssue.md)
    
*  **Photo Sync Issue**
   This feature helps user compare their profile photo from AAD to SPO user profile.
    
  <IMG src=.\assets\CheckUserProfilePhoto.JPG>
    
[More details](https://github.com/abrcheng/SharePointOnlineQuickAssist/blob/main/Documents/UserPhotoSync.md)
    
* **User Info list sync issue**,
	User/Group mail haven’t been synced to user info list caused the mail can’t be send to User/Group in user alert, workflow 
	User/Group display name updated but haven’t synced to user info list cause mismatch issue 
	User’ phone number haven’t been synced to user info list cause mismatch issue 
	User’ job title haven’t been synced to user info list cause mismatch issue 
![image](https://user-images.githubusercontent.com/21354416/144986960-e4befdd6-b9d6-40a0-bb54-fc90ca8d0d70.png)
![image](https://user-images.githubusercontent.com/21354416/144987002-085d0652-f243-4c29-84b6-94452d3afdee.png)
[More details](https://github.com/abrcheng/SharePointOnlineQuickAssist/blob/main/Documents/UserInfoSyncIssue.md)
*  **OneDrive library sync issue**,
	 OneDrive sync button can't be found.
	 Library synced as read only 
	 ![image](https://user-images.githubusercontent.com/21354416/144987185-d18e2e24-5b35-4436-ba3f-002dd95819c7.png)
         ![image](https://user-images.githubusercontent.com/21354416/144987204-83e0eb28-9f10-4ec4-858a-cc31c78d0d32.png)
[More details](https://github.com/abrcheng/SharePointOnlineQuickAssist/blob/main/Documents/OneDriveLockIcon.md)

*  **Missing New/Display/Edit forms issue**
	  ![image](https://user-images.githubusercontent.com/21354416/147523270-bef520ae-e487-4414-a449-6348b31efe82.png)
[More details](https://github.com/abrcheng/SharePointOnlineQuickAssist/blob/main/Documents/MissingForms.md)

*  **Uneditable wiki page** This feature helps to detect layout issue which could cause a classic wiki page being uneditable.
	   
	  ![image](https://user-images.githubusercontent.com/102142347/162903702-ed3ed028-6701-49b9-ab32-a88c14cb3480.png)

	  [More details](https://github.com/abrcheng/SharePointOnlineQuickAssist/blob/main/Documents/UneditableClassicWiki.md)

**Please click the link [Deployment Approaches](https://github.com/abrcheng/SharePointOnlineQuickAssist/blob/main/Documents/Installation/ReadMe.md) for checking deployment steps**

## The following steps show you how to customise and or contribute to this tool,
	
To build and start using these projects, you'll need to clone and build the projects.

Clone this repository by executing the following command in your console:

```shell
git clone https://github.com/abrcheng/SharePointOnlineQuickAssist.git
```

Navigate to the cloned repository folder which should be the same as the repository name:

```shell
cd SharePointOnlineQuickAssist
```

To access the webpart use the following commands.

```shell
cd SPFX/SPOQA
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
* [Microsoft Privacy Statement – Microsoft privacy](https://privacy.microsoft.com/en-us/privacystatement) 
