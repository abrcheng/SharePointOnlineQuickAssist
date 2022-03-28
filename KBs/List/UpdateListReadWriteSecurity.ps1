#Set Parameters
$SiteURL = "https://xxxxx.sharepoint.com/sites/xxxxx"
$ListName = "xxxxxxx"
  
#Connect to SharePoint Online site
Connect-PnPOnline -Url $SiteURL  -Interactive
 
#Get the List
$List = Get-PnPList $ListName -Includes ReadSecurity
 
#Set List Item-Security
$List.ReadSecurity = 1 
$List.WriteSecurity = 1  
$List.Update()
Invoke-PnPQuery
