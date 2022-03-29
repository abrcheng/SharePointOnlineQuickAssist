
This script used for uploading user pictures to SharePoint online user profile directly.

1. Start PowerShell as administrator and install the SharePoint PNP Module

	a. Install-Module -Name "PNP.PowerShell"
	
2. Prepare all users' pictures with the file name same as the users' UPN in local folder (e.g. C:\Photos\Photos),


3. Export/Save the affected user list to UsersListFile.txt

![image](https://user-images.githubusercontent.com/21354416/151517552-413b9ce5-7dc6-4fe5-be48-d7a98d241638.png)


4. Save the UploadUserPictureToSPO.ps1 to local drive
5. Run it as below 
	
	.\UploadUserPictureToSPO.ps1 -**usersListFile** ".\UsersListFile.txt" -**mySiteHostSiteUrl** https://chengc-my.sharepoint.com -**photoPath** C:\Photos\Photos 
![image](https://user-images.githubusercontent.com/21354416/151515934-0579cdb1-f2e9-4842-9042-20c5bf5c99fa.png)

a. **usersListFile** is the user list file name

b. **mySiteHostSiteUrl** is your tenant’s my site host URL

c. **phtotoPath** is the folder which stored the prepared users' pictures


6. Check the FailedUsers.txtand ErrorMessage.txtfor accounts which failed to be synced 
 






