1. Installed the SharePoint PNP and Azure AD PowerShell Module 

Install-Module -Name AzureAD

Uninstall-Module -Name "PNP.PowerShell"

Uninstall-Module -Name "SharePointPnPPowerShellOnline"

Install-Module -Name "SharePointPnPPowerShellOnline" -RequiredVersion 3.20.2004.0

3. If you need to update the Photo in Exchange online(EXO) as well, then install the EXO PowerShell according to https://docs.microsoft.com/en-us/powershell/exchange/connect-to-exchange-online-powershell?view=exchange-ps
4. Export/Save the affected user list to UsersListFile.txt

![image](https://user-images.githubusercontent.com/21354416/151517552-413b9ce5-7dc6-4fe5-be48-d7a98d241638.png)


3. Save the above script as SyncPhotoFromAADToSPOByGraphAPI.ps1
4. Run it as below 
	
	.\SyncPhotoFromAADToSPOByGraphAPI.ps1 -**usersListFile** ".\UsersListFile.txt" -**mySiteHostSiteUrl** https://chengc-my.sharepoint.com -**photoPath** C:\Photos\Photos -**updateExo** $false

![image](https://user-images.githubusercontent.com/21354416/151515934-0579cdb1-f2e9-4842-9042-20c5bf5c99fa.png)

a. **usersListFile** is the user list file name

b. **mySiteHostSiteUrl** is your tenant’s my site host URL

c. **phtotoPath** is the temp folder for storing the photo which download from AAD

d. **updateExo** switch specifies whether to update the photo via EXO command Set-UserPhoto in Exchange Online or not

5. Check the FailedUsers.txtand ErrorMessage.txtfor accounts which failed to be synced 
 
Please note the above script is based on the result of command “Get-AzureADUserThumbnailPhoto -ObjectId $user”, if that command can’t get the photo from AAD, then the script will can’t sync it either.





