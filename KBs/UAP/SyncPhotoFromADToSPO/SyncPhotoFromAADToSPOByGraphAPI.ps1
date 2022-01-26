param(  
  [Parameter(mandatory=$true)]
  [string]$mySiteHostSiteUrl = "https://chengc-my.sharepoint.com",
  [Parameter(mandatory=$true)]
  [string]$usersListFile = "D:\files\photos\UsersListFile.txt",
  [Parameter(mandatory=$true)]
  [string]$photoPath = "D:\files\photos\Photos"
)

# connect AAD for get the users's ThumbnailPhoto by commmand Get-AzureADUserThumbnailPhoto -ObjectId "CVDu@chengc.onmicrosoft.com"
Connect-PnPOnline -Scopes "User.Read","User.ReadBasic.All"
$accessToken =Get-PnPAccessToken

# Connect to the for uploading photos and update user profile
$adminSiteUrl = $mySiteHostSiteUrl.Replace("-my", "-admin")
$photoFolderUrl = "/User Photos/Profile Pictures/"
$users = Get-Content $usersListFile
Connect-PnPOnline $adminSiteUrl  -UseWebLogin
$adminCtx = Get-PnPContext
$null = Get-PnPUserProfileProperty -Account $users[0]

Connect-PnPOnline $mySiteHostSiteUrl -UseWebLogin
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
        #$null = Get-AzureADUserThumbnailPhoto -ObjectId $user -FilePath $photoPath -ErrorAction Stop
        $apiUrl = "https://graph.microsoft.com/v1.0/users/$user/photo/`$value"
        Invoke-RestMethod -Headers @{Authorization = "Bearer $accessToken"} -Uri $apiUrl -Method Get -ContentType image/jpeg -OutFile $photoPath\$user.jpg

        $file = Get-ChildItem -Path $photoPath -Filter "$user*"
        if($file.FullName -ne $empty)
        {
            $baseFilename = $user.Replace(".", "_").Replace("@","_")
            GeneratetThumbnail -filePath $file.FullName -baseFilename $baseFilename

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
             $errorMessage = "Failed to get the user's phtot when processing $($user) "
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