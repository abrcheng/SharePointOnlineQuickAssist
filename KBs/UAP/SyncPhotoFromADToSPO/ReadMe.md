1. Installed the SharePoint PNP and Azure AD PowerShell Module 
2. Export/Save the affected user list to UsersListFile.txt

![image](https://user-images.githubusercontent.com/21354416/151100496-3493c7df-0f19-4087-8173-c3a46d5db5d8.png)

3. Save the above script as SyncPhotoFromAADToSPO.ps1
4. Run it as below 
	
	.\SyncPhotoFromAADToSPO.ps1 -**usersListFile** ".\UsersListFile.txt" -**mySiteHostSiteUrl** https://chengc-my.sharepoint.com-photoPath C:\Photos\Photos
	
a. usersListFile is the user list file name

b. mySiteHostSiteUrl is your tenant’s my site host URL

c. phtotoPath is the temp folder for storing the photo which download from AAD

1. Check the FailedUsers.txtand ErrorMessage.txtfor accounts which failed to be synced 
 
Please note the above script is based on the result of command “Get-AzureADUserThumbnailPhoto -ObjectId $user”, if that command can’t get the photo from AAD, then the script will can’t sync it either.

If there is a error message as below, but the photo can be got by by Graph API, then please use **SyncPhotoFromAADToSPOByGraphAPI.ps1**,
	Get-AzureADUserThumbnailPhoto : Error occurred while executing GetAzureADUserThumbnailPhoto
	Code: Request_ResourceNotFound
	Message: Resource 'thumbnailPhoto' does not exist or one of its queried reference-property objects are not
	present.
![image](https://user-images.githubusercontent.com/21354416/151100735-52c86402-46e4-4ba5-b90b-56aeb9fba64d.png)

Please note **SyncPhotoFromAADToSPOByGraphAPI.ps1** need to install PNP with below version,

Uninstall-Module -Name "PNP.PowerShell"

Uninstall-Module -Name "SharePointPnPPowerShellOnline"

Install-Module -Name "SharePointPnPPowerShellOnline" -RequiredVersion 3.20.2004.0
