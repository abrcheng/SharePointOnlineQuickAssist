# 21V auth need to configure by https://pnp.github.io/powershell/articles/authentication.html#authentication-to-gcc-or-national-cloud-environments 
# https://pnp.github.io/powershell/articles/authentication.html#special-instructions-for-gcc-or-national-cloud-environments 
 
 
#OneDrive mysite host URL 
$mySiteHostSiteUrl = "https://abrcheng-my.sharepoint.cn"
#user UPNs
$usersListFile = "c:\files\photos\UsersListFile.txt"
#Photo folder in local device (blank)
$photoPath = "c:\files\photos\Photos"
#Need update for EXO or not
$updateExo = $false 

# App client ID
$ClientID = "c4faf394-90e2-4529-a976-845afae3476e"
$TenantID = "9e658c2c-aa86-44db-9a3f-fb6aa9033ee9"
 

# connect AAD for get the users's ThumbnailPhoto by commmand Get-AzureADUserThumbnailPhoto -ObjectId "CVDu@chengc.onmicrosoft.com"


# https://pnp.github.io/powershell/articles/authentication.html#authentication-to-gcc-or-national-cloud-environments


#You must register your own applications and use them when connecting to Microsoft Graph. https://learn.microsoft.com/en-us/entra/identity-platform/quickstart-register-app 
Connect-MgGraph -Environment China -Scopes "User.Read.All", "Group.ReadWrite.All" -ClientId $ClientID -TenantId $TenantID 

# Connect to exchange online
if($updateExo)
{
    Connect-ExchangeOnline 
}

# Connect to the for uploading photos and update user profile
$adminSiteUrl = $mySiteHostSiteUrl.Replace("-my", "-admin")
$photoFolderUrl = "/User Photos/Profile Pictures/"
$users = Get-Content $usersListFile
# Update: Clear cached credential to connect to site
Connect-PnPOnline $adminSiteUrl -Interactive -ClientId $ClientID -Tenant $TenantID -AzureEnvironment China
$adminCtx = Get-PnPContext
$null = Get-PnPUserProfileProperty -Account $users[0] # load Microsoft.SharePoint.Client.UserProfiles

Connect-PnPOnline  $mySiteHostSiteUrl -Interactive -ClientId $ClientID -Tenant $TenantID -AzureEnvironment China
$ctx = Get-PnPContext

$peopleManager = New-Object Microsoft.SharePoint.Client.UserProfiles.PeopleManager($adminCtx)

function GeneratetThumbnail($filePath, $baseFilename)
{
  $full = [System.Drawing.Image]::FromFile($filePath)

  $midThumb = $full.GetThumbnailImage(72, 72, $null, [intptr]::Zero)
  $midPath = "$photoPath\$baseFilename"+"_MThumb.jpg"
  $midThumb.Save($midPath);

  $smallThumb = $full.GetThumbnailImage(48, 48, $null, [intptr]::Zero)
  $smallPath = "$photoPath\$baseFilename"+"_SThumb.jpg"
  $smallThumb.Save($smallPath)

  $largeThumb = $full.GetThumbnailImage(300, 300, $null, [intptr]::Zero)
  $largePath = "$photoPath\$baseFilename"+"_LThumb.jpg"
  $largeThumb.Save($largePath)

  $full.Dispose()
  $midThumb.Dispose()
  $midThumb.Dispose()
  $largeThumb.Dispose()    
} 

foreach($user in $users)
{
    try
    {
        Write-Host "Processing $user"        
        #$null = Get-EntraUserThumbnailPhoto -ObjectId $user -FilePath $photoPath -ErrorAction Stop
        #$apiUrl = "https://microsoftgraph.chinacloudapi.cn/v1.0/users/$user/photo/`$value"
        #Invoke-RestMethod -Headers @{Authorization = "Bearer $accessToken"} -Uri $apiUrl -Method Get -ContentType image/jpeg -OutFile $photoPath\$user.jpg
        $UserID = Get-MgUser -Filter "startsWith(UserPrincipalName, '$user')" -Top 1|select ID
        Get-MgUserPhotoContent -UserId $userID.Id -OutFile $photoPath\$user.jpg

        $file = Get-ChildItem -Path $photoPath -Filter "$user*"
        if($file.FullName -ne $empty)
        {
            $baseFilename = $user.Replace(".", "_").Replace("@","_")
            GeneratetThumbnail -filePath $file.FullName -baseFilename $baseFilename

            if($updateExo)
            {
                Set-UserPhoto -Identity $user -PictureData ([System.IO.File]::ReadAllBytes("$($photoPath)\$($baseFilename)_SThumb.jpg")) -Confirm:$false
            }
            
            $thumbnails = Get-ChildItem -Path $photoPath -Filter "$baseFilename*"
            foreach($thumbnail in $thumbnails)
            {       
                Write-Host "Uploading $($thumbnail.FullName)......"         
                $null = Add-PnPFile -Path $thumbnail.FullName -Folder $photoFolderUrl -ErrorAction Stop            
            }

            $fullPictureUrl = $mySiteHostSiteUrl + $photoFolderUrl + $baseFilename + "_MThumb.jpg"
            $uploadedPicture = Get-PnPFile $fullPictureUrl
            if($uploadedPicture.Length -gt 0)
            {
                $peopleManager.SetSingleValueProfileProperty("i:0#.f|membership|" + $user, "PictureUrl", $fullPictureUrl)
                $peopleManager.SetSingleValueProfileProperty("i:0#.f|membership|" + $user, "SPS-PicturePlaceholderState", 0)
                $peopleManager.SetSingleValueProfileProperty("i:0#.f|membership|" + $user, "SPS-PictureExchangeSyncState", 1)
                $adminCtx.ExecuteQuery()
            }
        }
        else
        {
             $errorMessage = "Failed to get the user's photo when processing $($user) "
             Write-Host $errorMessage
             $errorMessage >>"ErrorMessage.txt"  
             $user  >> "FailedUsers.txt"
        }
        
    }
    catch [System.Exception] 
    {
          $exceptionMessage = $_.Exception.ToString()
          $errorMessage = "Get error message $exceptionMessage when processing $($user)"
          Write-Host $errorMessage
          $errorMessage >>"ErrorMessage.txt"  
          $user >> "FailedUsers.txt"
    }     
}

Write-Host "Done!"  -ForegroundColor Green
