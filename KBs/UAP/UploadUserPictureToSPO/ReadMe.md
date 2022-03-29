
This script used for uploading user pictures from local drive to SharePoint online user profile directly.

1. Start PowerShell as administrator and install the SharePoint PNP Module

	a. Install-Module -Name "PNP.PowerShell"
	
2. Prepare all users' pictures with the file name same as the users' UPN in local folder (e.g. C:\Photos\Photos),
![image](https://user-images.githubusercontent.com/21354416/160579725-d265f2fa-f01c-48fd-9e27-21620914ddd5.png)


3. Export/Save the affected user list to UsersListFile.txt
![image](https://user-images.githubusercontent.com/21354416/160580116-632a35d6-c0ea-4da3-a67b-2e363704d2dc.png)


4. Save the UploadUserPictureToSPO.ps1 to local drive
5. Run it as below 
	
	.\UploadUserPictureToSPO.ps1 -**usersListFile** ".\UsersListFile.txt" -**mySiteHostSiteUrl** https://chengc-my.sharepoint.com -**photoPath** C:\Photos\Photos 

![image](https://user-images.githubusercontent.com/21354416/160579874-0c30b044-878e-4957-99a9-86beea5b4ebf.png)

a. **usersListFile** is the user list file name

b. **mySiteHostSiteUrl** is your tenant’s my site host URL

c. **phtotoPath** is the folder which stored the prepared users' pictures


6. Check the FailedUsers.txtand ErrorMessage.txtfor accounts which failed to be synced 
 






