################################################################################
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
# THE SOFTWARE.
#
# Filename: GetDocIDReport.ps1
# Description: This script will generate a report for document ID assignment in the site
# Output: 
#    1. xxxxx_MissDocIdLog.txt, if there is any documents which missed the document ID will be logged
#    2. xxxxx_DocIDSummaryReport.csv, Document ID assignment summary report in list level
#    3. xxxxx_GetDocIDReportErrorLog.log, if there is any error message when executing the script will be logged
# Paramters:
#    1. siteUrl, the site URL(e.g. https://chengc.sharepoint.com/sites/abc) whicn need to be processed 
##################################################################################
[Cmdletbinding()]
Param (
    [Parameter(mandatory=$true)]
    [String] $siteUrl
)
$startTime = $([System.DateTime]::Now)
Write-Host "start time:$startTime" 

$global = [PSCustomObject]@{
            ErrorCount = 0          
            DocumentCount = 0
            ItemsCount = 0
            LibraryCount = 0
            AssignedDocIDCount = 0
            FolderCount=0
            }

$skipLibs = @("PreservationHoldLibrary")
function LogError($message){
   $global.ErrorCount++
   Write-Host "$message" -ForegroundColor Red
   $message | out-file $errorLogPath -Encoding ascii -Force -Append  
}

function ProcessLibrary([Microsoft.SharePoint.Client.List] $lib, [Microsoft.SharePoint.Client.ClientContext] $clientContext) # Process Library
{
    $libRootFolderUrl = $lib.RootFolder.ServerRelativeUrl
    foreach($skipLib in $skipLibs)
    {
        if($libRootFolderUrl.EndsWith($skipLib))
        {
              Write-Host "Skip library $libRootFolderUrl ..."
              return
        }
    }

    Write-Host "Processing library $libRootFolderUrl ..."
    $global.LibraryCount++ 
    $libDoucmentCount=0
    $libItemCount=0
    $libFolderCount=0
    $libDocAssignedDocIDCount=0

    $listItePosition = New-Object Microsoft.SharePoint.Client.ListItemCollectionPosition
    $listItePosition.PagingInfo = ""
    $caml =  "<View Scope='RecursiveAll'>  
                     <Query></Query>                 
                     <ViewFields><FieldRef Name='FileRef' /><FieldRef Name='ID' /><FieldRef Name='_dlc_DocId' /><FieldRef Name='FSObjType' /></ViewFields>    
                     <RowLimit Paged='TRUE'>1000</RowLimit> 
              </View>"
    $oQuery = New-Object Microsoft.SharePoint.Client.CamlQuery 
    $oQuery.ViewXml = $caml
    do
    {
       $oQuery.ListItemCollectionPosition = $listItePosition       
       $items = $lib.GetItems($oQuery)
       try
       {
            $clientContext.Load($items)
            $clientContext.ExecuteQuery()
            [System.Threading.Thread]::Sleep(1000) # In order to avoid throttling sleep 1 seconds for very query 
       }
       catch [System.Exception]
       {
          $errorMessge = "Get error $($_.Exception.ToString()) when procesing library $($lib.RootFolder.ServerRelativeUrl)"
          Write-Host $errorMessge -ForegroundColor Red

          # Wait for 10 seconds and retry again 
          [System.Threading.Thread]::Sleep(10*1000)
          $clientContext.Load($items)
          $clientContext.ExecuteQuery()
          Write-Host "Retry successful, this error can be ignored" -ForegroundColor Green  
       }
       $itemCount = $items.Count
       Write-Host "$([System.DateTime]::Now) Get $itemCount items in this batch"

       foreach($item in $items)
       {
          $global.ItemsCount++ 
          $libItemCount++
          $filePath = $item["FileRef"]
          if($item.FileSystemObjectType -eq [Microsoft.SharePoint.Client.FileSystemObjectType]::File)
          {
             $libDoucmentCount++  
             $global.DocumentCount++   
             if($item["_dlc_DocId"] -ne $null)  
             {
                $global.AssignedDocIDCount++
                $libDocAssignedDocIDCount++
             }
             else
             {
                 $filePath >> $missDocIdLogPath
             }
          }      
          else
          {
            $global.FolderCount++
            $libFolderCount++
          }
       }   

       $listItePosition = $items.ListItemCollectionPosition

    } While($listItePosition -ne $null)

    $listSummary = [PSCustomObject]@{
            Library =  $libRootFolderUrl          
            ItemCount = $libItemCount
            FolderCount = $libFolderCount
            DoucmentCount = $libDoucmentCount
            AssignedDocIDCount = $libDocAssignedDocIDCount
            MissedDocIDCount = $libDoucmentCount -$libDocAssignedDocIDCount 
            }

    $null=$libLevelSummaryList.Add($listSummary)
}

function ProcessWeb([Microsoft.SharePoint.Client.Web] $web, [Microsoft.SharePoint.Client.ClientContext] $clientContext)
{
    # Processs all libries of this web
    Write-Host "Processing site/subsite $($web.Url) ..."
    $lists = $web.Lists
    $clientContext.Load($lists)
    $clientContext.ExecuteQuery()
    foreach($list in $lists)
    {
         try
         {
            # Only process the library
            if($list.BaseType -eq [Microsoft.SharePoint.Client.BaseType]::DocumentLibrary -and $list.Hidden -eq $false)
            {   
                $clientContext.Load($list.RootFolder)
                $clientContext.ExecuteQuery()
                ProcessLibrary -lib $list -clientContext $clientContext
            }
         }
         catch [System.Exception]
         {             
             $errorMessge = "Get error $($_.Exception.ToString()) when procesing library $($list.DefaultViewUrl)"
             LogError -message $errorMessge 
         }
    }

    # Process all sub webs of this web
    $subWebs = $web.Webs
    $clientContext.Load($subWebs)
    $clientContext.ExecuteQuery()
    foreach($subWeb in $subWebs)
    {
        try
        {   
            ProcessWeb -web $subWeb -clientContext $clientContext
        }
        catch [System.Exception]
         {             
             $errorMessge = "Get error $($_.Exception.ToString()) when procesing sub web $($subWeb.Url)"
             LogError -message $errorMessge  
         }
    }    
}

Connect-PnPOnline -Url $siteUrl -UseWebLogin 
$clientContext = Get-PnPContext

$web = $clientContext.Web
$clientContext.Load($web)
$clientContext.ExecuteQuery()
$webTitle = $web.Title

$missDocIdLogPath = "$($webTitle)_MissDocIdLog.txt"
$summaryReportPath = "$($webTitle)_DocIDSummaryReport.csv"
$errorLogPath = ".\$($webTitle)GetDocIDReportErrorLog.log"

$libLevelSummaryList = [System.Collections.ArrayList]@()

ProcessWeb -web $web -clientContext $clientContext

$libLevelSummaryList | Export-Csv $summaryReportPath -NoTypeInformation
Write-Host "Processed $($global.LibraryCount) libraries,$($global.ItemsCount) items, $($global.FolderCount) folders, $($global.DocumentCount) documents, $($global.AssignedDocIDCount) documents have been assinged DocID" -ForegroundColor Green
if($global.DocumentCount - $global.AssignedDocIDCount -ne 0)
{
    Write-Host "$($global.DocumentCount - $global.AssignedDocIDCount) documents haven't been assigned document ID" -ForegroundColor Red
}
if($global.ErrorCount -lt 0)
{
   Write-Host "Hit $($global.ErrorCount) errors, please check the $($errorLogPath) for more detail error message" -ForegroundColor Red
}

$totalMinutes = [int]([System.DateTime]::Now -$startTime).TotalMinutes
Write-Host "End time:$([System.DateTime]::Now), duration is $totalMinutes minutes" 
Write-Host "Done." -ForegroundColor Green 