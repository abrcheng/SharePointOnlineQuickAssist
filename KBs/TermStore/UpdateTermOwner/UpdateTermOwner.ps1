Connect-PnPOnline https://chengc.sharepoint.com -UseWebLogin
$TermGroupName = "Site Collection - chengc.sharepoint.com"
$TermSetName= "SiteMM"
$TermSetOwner= "i:0#.f|membership|johna@chengc.onmicrosoft.com"
$processedCount = 0

$Ctx = Get-PnPContext
$TaxonomySession = [Microsoft.SharePoint.Client.Taxonomy.TaxonomySession]::GetTaxonomySession($Ctx)
$TaxonomySession.UpdateCache()
$Ctx.Load($TaxonomySession)

$TermStore = $TaxonomySession.GetDefaultSiteCollectionTermStore()
$Ctx.Load($TermStore)
$Ctx.ExecuteQuery()
 
$TermGroup = $TermStore.Groups.GetByName($TermGroupName)
$Ctx.Load($TermGroup)

#Get the termset
$TermSet = $TermGroup.TermSets.GetByName($TermSetName)
$Ctx.Load($TermSet)
$Ctx.ExecuteQuery()

$terms = $TermSet.GetAllTerms()
$Ctx.Load($terms)
$Ctx.ExecuteQuery()

foreach($term in $terms)
{
  Write-Host "Processing $($term.Name) ....."
  $term.Owner = $TermSetOwner
  $processedCount++
}

$TermStore.CommitAll()
$Ctx.ExecuteQuery()
Write-host "Processed $processedCount terms, Done" -ForegroundColor Green