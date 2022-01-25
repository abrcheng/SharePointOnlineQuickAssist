param( 
  [string]$adminSiteUrl = "https://chengc-admin.sharepoint.com"  
)

Connect-PnPOnline -Url $adminSiteUrl -UseWebLogin
$clientContext = Get-PnPContext
Connect-AzureAD
$allUsers = Get-AzureADUser -all $true
$userCount =0
foreach($user in $allUsers)
{
    if($user.UserPrincipalName.Contains("#EXT#") -ne $true)
    {
       Write-Host "Processing $($user.UserPrincipalName)"
       try
       {
           $Properties = Get-PnPUserProfileProperty -Account $user.UserPrincipalName
           if([System.String]::IsNullOrWhiteSpace($Properties.PictureUrl))
           {
                 Write-Host  $user.UserPrincipalName
                 $user.UserPrincipalName >>"UsersWithOutPictureInSPO.txt"
                 $userCount++
           }
       }
       catch [System.Exception] 
       {
          $errormessage = $_.Exception.ToString()
          "Get error message $errormessage when processing $($user.UserPrincipalName) \r\n" >>"ErrorMessage.txt"  
       }       
    }
}
 
Write-Host "Detected $userCount users without picture in the SPO user profile, please check UsersWithOutPictureInSPO.txt for detail"
Write-Host "Done" -ForegroundColor Green
