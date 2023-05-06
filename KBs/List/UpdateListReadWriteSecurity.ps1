#Set Parameters
$SiteURL = "https://mykademia.sharepoint.com/sites/mykademia"
$ListName = "mykademia"
  
#Connect to SharePoint Online site
Connect-PnPOnline -Url $SiteURL  -Interactive
 
#Get the List
$List = Get-PnPList $ListName -Includes ReadSecurity
 
#Set List Item-Security
$List.ReadSecurity = 1 
$List.WriteSecurity = 1  
$List.Update()
Invoke-PnPQuery
