$StartDate = "2019/01/24 9:00:00"
$EndDate = "2019/01/25 9:00:00"
$SiteURL = "https://chengc.sharepoint.com"

Connect-PnPOnline $SiteURL  

$DeletedCollection = Get-PnPRecyclebinItem -RowLimit 500 | ? { ($_.DeletedByEmail -eq 'UserEmailID') -and ($_.DeletedDate -ge $StartDate) -and ($_.DeletedDate -le $EndDate)}
                                                                              
 $ctx = Get-PnPContext
 

foreach($DeletedItem in $DeletedCollection)
{
    Try
    {
    Write-host "Initiating Restore for: "  $DeletedItem.Title
$RestoreURL = $SiteURL+"/_api/web/Recyclebin('$($DeletedItem.ID)')/restore()"
    Invoke-PnPSPRestMethod -Method post -Url $RestoreURL 

    Write-host "Restore Completed for: "  $DeletedItem.Title
    Write-Host -ForegroundColor Green "Success :: File is successfully restored"
    Write-host "Writing to log file.................____________________________________________ " $Successcount
    $Successcount++
    }
    Catch
    {
     Write-host "Exception for: "  $DeletedItem.Title
     $errormessage = "Catch Error:" + $_.Exception.Message
     Write-host "Writing to log file.................____________________________________________ " $FailureCount
    $FailureCount++
    Write-Host -ForegroundColor Red "Catch :: There is an error while restoring the Document"  $_.Exception.Message
    Continue
    }
}

Write-Host -ForegroundColor DarkGreen "Success Count: " + $Successcount
Write-Host -ForegroundColor DarkRed "Failure Count: " + $FailureCount    
