1. Installed the SharePoint PNP and Azure AD PowerShell Module (please refer these steps which I sent to you on 20th July 2021)
2. Export the user list which need to be synced (I have provided the script to you)
 
· Steps for running the above script
3. Save the above script as SyncPhotoFromAADToSPO.ps1
4. Run it as below 
	
	
	.\SyncPhotoFromAADToSPO.ps1 -usersListFile ".\UsersListFile.txt" -mySiteHostSiteUrl https://chengc-my.sharepoint.com-photoPath C:\Photos\Photos
a. usersListFile is the user list file name
			
			
b. mySiteHostSiteUrl is your tenant’s my site host URL
c. phtotoPath is the temp folder for storing the photo which download from AAD
1. Check the FailedUsers.txtand ErrorMessage.txtfor accounts which failed to be synced 
	
	 
	
	
 
Please note the above script is based on the result of command “Get-AzureADUserThumbnailPhoto -ObjectId$user”, if that command can’t get the photo from AAD, then the script will can’t sync it either.
