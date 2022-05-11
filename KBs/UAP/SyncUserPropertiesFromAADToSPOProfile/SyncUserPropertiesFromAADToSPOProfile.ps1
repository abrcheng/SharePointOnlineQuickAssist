# This script is used for syncing user properties from ADD to SharePoint Online user profile 
param(
  [Parameter(mandatory=$true)]
  [string]$AdminSiteURL = "https://chengc-admin.sharepoint.com", # SPO admin site URL
  [bool]$IsChinaCloud = $false, # IsChinaCloud indicates the AzureEnvironment
  [bool]$SyncManager= $false, # will not sync manage by default
  [bool]$SyncWorkMail=$false
)

$ADProperties = @("Department","Department","GivenName","Surname","DisplayName","telephoneNumber","JobTitle")
$SPOProperies = @("Department","SPS-Department","FirstName","LastName","PreferredName","WorkPhone","SPS-JobTitle")

# AzureEnvironmentName for Connect-AzureAD: AzureCloud,AzureChinaCloud,AzureUSGovernment,AzureGermanyCloud
$EnvironmentForAAD = "AzureCloud"

if($IsChinaCloud)
{
   $EnvironmentForAAD = "AzureChinaCloud"  
}

Connect-AzureAD -AzureEnvironmentName  $EnvironmentForAAD  | Out-Null
Connect-PnPOnline -Url $AdminSiteURL -UseWebLogin | Out-Null
$AllUsers = $null

if($SyncWorkMail) # proxyAddresses => WorkEmail	
{
    $ADProperties.Add("Mail")
    $SPOProperies.Add("WorkEmail")
}

if($SyncManager)
{
    $ADProperties.Add("Manager")
    $SPOProperies.Add("Manager")
    $AllUsers = Get-AzureADUser -All:$True -Filter "UserType eq 'Member'" | select *,@{n="Manager";e={(Get-AzureADUser -ObjectId (Get-AzureADUserManager -ObjectId $_.ObjectId).ObjectId).UserPrincipalName}}
}
else
{
    $AllUsers = Get-AzureADUser -All:$True -Filter "UserType eq 'Member'"
}

Write-host "Queried $($AllUsers.Count) users." 
$UpdateCount =0
 forEach($User in $AllUsers)
 {
     $updated =$false
     try
     {
         Write-host "Processing $($User.UserPrincipalName) ......" 
         $UserAccount = "i:0#.f|membership|$($User.UserPrincipalName)"
         $UserProfile = Get-PnPUserProfileProperty -Account $UserAccount
         if([System.String]::IsNullOrEmpty($UserProfile.AccountName))
         {
              write-host "Can't find the SPO profile for account $UserAccount, skip it" -ForegroundColor Yellow
              continue
         }

         for($index=0; $index -lt $SPOProperies.Length; $index++)
         {
             $SPOPropertyName = $SPOProperies[$index]
             $AADPropertyName = $ADProperties[$index]
         
             $SPOPropertyValue = $UserProfile.UserProfileProperties[$SPOPropertyName]
             $AADPropertyValue = $User.$AADPropertyName
             
             if([System.String]::IsNullOrEmpty($AADPropertyValue))
             {
                $AADPropertyValue = ""
             }

             if([System.String]::IsNullOrEmpty($SPOPropertyValue))
             {
                $SPOPropertyValue = ""
             }

             if($SPOPropertyValue -ne $AADPropertyValue)
             {
                Write-Host "Detected mismatch: $SPOPropertyName in SPO is $SPOPropertyValue, $AADPropertyName in AAD is $AADPropertyValue, will update"
                Set-PnPUserProfileProperty -Account $UserAccount -PropertyName $SPOPropertyName -Value $AADPropertyValue
                $updated = $true
             }
         }

         if($updated)
         {
             $UpdateCount++
         }
     }
     catch [System.Exception]
     {
         $errormessage = $_.Exception.ToString()
         write-host $errormessage -ForegroundColor Red
     }
 }

 Write-Host "Updated $UpdateCount users." -ForegroundColor Green