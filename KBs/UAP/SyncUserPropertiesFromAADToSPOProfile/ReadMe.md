 This script is used for syncing user properties from ADD to SharePoint Online user profile.
 
 It will sync below propeties by default,
	"Department","GivenName","Surname","DisplayName","telephoneNumber","JobTitle".

And if you need to sync WorkMail and Manager, then need to add the addtional parameters (-SyncWorkMail $true -SyncManager $true)
**It neeed to be run by SharePoint Online tenant admin.**


1. Start PowerShell as administrator and install the SharePoint PNP and Azure AD PowerShell Module

	a. Install-Module -Name AzureAD
	
	b. Uninstall-Module -Name "SharePointPnPPowerShellOnline"
	
	c. Install-Module -Name "PnP.PowerShell"
	
	d. If the computer doesn't allow to run PowerShell script, allow it by running Set-ExecutionPolicy according to https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.security/set-executionpolicy?view=powershell-7.2

2. Download the script **SyncUserPropertiesFromAADToSPOProfile.ps1** to local drive

3. Run the script as below for China 21v enviroment (please repalce the place holder contonso), 

		.\SyncUserPropertiesFromAADToSPOProfile.ps1 -AdminSiteURL https://contonso-admin.sharepoint.cn -IsChinaCloud $true
![image](https://user-images.githubusercontent.com/21354416/167838200-e946942d-2306-48d8-8c98-80a9d665384f.png)

		
4. Run the script as below for Global enviroment (please repalce the place holder contonso), 

		.\SyncUserPropertiesFromAADToSPOProfile.ps1 -AdminSiteURL https://contonso-admin.sharepoint.com  






