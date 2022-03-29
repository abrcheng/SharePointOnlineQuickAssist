1. Start PowerShell as administrator and install the SharePoint PNP and Azure AD PowerShell Module

	a. Install-Module -Name AzureAD

	b. Uninstall-Module -Name "PNP.PowerShell"

	c. Uninstall-Module -Name "SharePointPnPPowerShellOnline"

	d. Install-Module -Name "SharePointPnPPowerShellOnline" -RequiredVersion 3.20.2004.0

2. If you need to update the Photo in Exchange online(EXO) as well, then install the EXO PowerShell according to https://docs.microsoft.com/en-us/powershell/exchange/connect-to-exchange-online-powershell?view=exchange-ps
3. Export/Save the affected user list to UsersListFile.txt

![image](https://user-images.githubusercontent.com/21354416/151517552-413b9ce5-7dc6-4fe5-be48-d7a98d241638.png)


4. Save the SyncPhotoFromAADToSPOByGraphAPI.ps1 to local drive
5. Run it as below 
	
	.\SyncPhotoFromAADToSPOByGraphAPI.ps1 -**usersListFile** ".\UsersListFile.txt" -**mySiteHostSiteUrl** https://chengc-my.sharepoint.com -**photoPath** C:\Photos\Photos -**updateExo** $false

![image](https://user-images.githubusercontent.com/21354416/151515934-0579cdb1-f2e9-4842-9042-20c5bf5c99fa.png)

a. **usersListFile** is the user list file name

b. **mySiteHostSiteUrl** is your tenant’s my site host URL

c. **phtotoPath** is the temp folder for storing the photo which download from AAD

d. **updateExo** switch specifies whether to update the photo via EXO command Set-UserPhoto in Exchange Online or not

6. Check the FailedUsers.txtand ErrorMessage.txtfor accounts which failed to be synced 
 
Please note the above script is based on the result of command “Get-AzureADUserThumbnailPhoto -ObjectId $user”, if that command can’t get the photo from AAD, then the script will can’t sync it either.





