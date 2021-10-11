################################################################################
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
# THE SOFTWARE.
#
# Filename: TouchDocumentIds.ps1
# Description: Performs an update on the documents triggering the event receiver to assign document ids.
# Output: 
#    1. TouchDocumentIdsErrorMessage.txt, if there is any error message when executing the script will be logged
#    2. TouchDocumentIdsFailedFiles.txt, if there is any files failed to be processed will be logged
# Paramters:
#    1. siteUrl, the site URL(e.g. https://chengc.sharepoint.com/sites/abc) whicn need to be processed 
#    2. listTitle, the title of the list which need to be processed
#    3. batchSize, number of the changes need to be submited in a batch
# Please Note: This script will trigger the workflow or PowerAutomate flow which associated with the library 
##################################################################################
[Cmdletbinding()]
Param (
    [Parameter(mandatory=$true)] # SiteUrl e.g. https://chengc.sharepoint.com/sites/abc
    [String] $siteUrl, 
    [Parameter(mandatory=$true)]
    [String]$listTitle, # Library title, e.g. Documents
    [Parameter(mandatory=$true)]
    [int]$batchSize =10
    )
 
Connect-PnPOnline $siteUrl 
[Microsoft.SharePoint.Client.ClientContext] $clientContext = Get-PNPContext

# Set time out to 10 minutes
$clientContext.RequestTimeout = 1000*60*10; 
$startTime = $([System.DateTime]::Now)
Write-Host "start time:$($startTime)" 
$listItePosition = New-Object Microsoft.SharePoint.Client.ListItemCollectionPosition
$listItePosition.PagingInfo = ""
$caml = "<View Scope='RecursiveAll'>  
                      <Query></Query>  
                      <ViewFields> 
                        <FieldRef Name='ID' /> 
                        <FieldRef Name='FileRef'/>  
                        <FieldRef Name='FSObjType' />  
                        <FieldRef Name='_dlc_DocId' />                        
                      </ViewFields>    
                      <OrderBy><FieldRef Name='ID' Ascending='True'/></OrderBy>                   
                      <RowLimit Paged='TRUE'>1000</RowLimit>  
                    </View>"
         
$oQuery = New-Object Microsoft.SharePoint.Client.CamlQuery 
$oQuery.ViewXml = $caml   
$totalProcessedCount = 0
$list = $clientContext.Web.Lists.GetByTitle($listTitle)

$assigned =0
$errorCount=0
do
   {      
      $currentBatchUnSumbitCount =0
      $oQuery.ListItemCollectionPosition = $listItePosition
      $items = $list.GetItems($oQuery)
      $clientContext.Load($items)
      $clientContext.ExecuteQuery()
      $itemCount = $items.Count
      $totalProcessedCount = $totalProcessedCount + $itemCount
      Write-Host "Get $itemCount items from the list/library $libName"
      for($index=$itemCount -1; $index -ge 0; $index--)
      {    
           $currentItem = $items[$index]       
           $url = $currentItem["FileRef"]       
           if($currentItem.FileSystemObjectType -eq [Microsoft.SharePoint.Client.FileSystemObjectType]::File -and $currentItem["_dlc_DocId"] -eq $null)
           {  
               Write-Host "Processing $url"
               $currentItem.SystemUpdate()
               $assigned++
               $currentBatchUnSumbitCount++
             
           }
          else
           {
                Write-Host "Skip $url as it is not a file or document ID has been assigned"    
           }
           
           if($currentBatchUnSumbitCount -eq $batchSize -or ($index -eq 0 -and $currentBatchUnSumbitCount -gt 0))
               {     
                   try
                   {
                       $clientContext.ExecuteQuery()                                
                   }
                   catch [System.Exception] 
                   {
                      $exceptionMessage = $_.Exception.ToString()
                      $errorMessage = "Get error message $exceptionMessage when processing $url"
                      Write-Host $errorMessage -ForegroundColor Red
                      $errorMessage >>"TouchDocumentIdsErrorMessage.txt"  
                      $url >> "TouchDocumentIdsFailedFiles.txt"
                      $errorCount++
                        
                      # Wait for 10 seconds and retry again 
                      [System.Threading.Thread]::Sleep(10*000)
                      $clientContext.ExecuteQuery()  
                      $ignoreMessage = "Retry successful, this error can be ignored"
                      $ignoreMessage >>"TouchDocumentIdsErrorMessage.txt"  
                      Write-Host $ignoreMessage   -ForegroundColor Green                    
                  }

                  $currentBatchUnSumbitCount =0
              }                       
      }                  
      
      $listItePosition = $items.ListItemCollectionPosition
   } While($listItePosition -ne $null) 
 
Write-Host "end time:$([System.DateTime]::Now)" 
Write-Host "Updated $assigned documents with $errorCount errors"
$totalMinutes = [int]([System.DateTime]::Now -$startTime).TotalMinutes
Write-Host "End time:$([System.DateTime]::Now), duration is $totalMinutes minutes" 
Write-Host "Done." -ForegroundColor Green 
