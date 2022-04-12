**[Symptoms]**

When the user enter edit mode on a classic wiki page, the ribbon menu may gray out and the page became uneditable.

**[Cause]**

Classic wiki page holds its layout data in an OOTB field as HTML. If the field contains invalid value or the layout table in the HTML does not match the declaration, it could lead to this issue.

**[Solution]**

The layout data is being saved in OOTB field named "WikiField". Below is a sample checking the value using PnP powershell cmdlet. The sample page here has "Three column with header and footer" layout.
```powershell
    Connect-PnPOnline https://xxxxxx.sharepoint.com/sites/xxxxxxx
    $pages = Get-PnPListItem -List sitepages
    # List up all pages and shows each page's Id and Filename 
    $pages | Select-Object id,@{label="Filename";expression={$_.FieldValues.FileLeafRef}}
    # Use target page's Id to retreive WikiField value
    ($page | ? Id -eq 11111).FieldValues.WikiField
    <div class="ExternalClass0ABFFA5383FD485EA2B0F7885197795D"><table id="layoutsTable" style="width&#58;100%;"><tbody><tr style="vertical-align&#58;top;"><td colspan="3"><div class="ms-rte-layoutszone-outer" style="width&#58;100%;"><div class="ms-rte-layoutszone-inner"></div>&#160;</div></td></tr><tr style="vertical-align&#58;top;"><td style="width&#58;33.3%;"><div class="ms-rte-layoutszone-outer" style="width&#58;100%;"><div class="ms-rte-layoutszone-inner"><p><br></p></div></div></td><td class="ms-wiki-columnSpacing" style="width&#58;33.3%;"><div class="ms-rte-layoutszone-outer" style="width&#58;100%;"><div class="ms-rte-layoutszone-inner"></div>&#160;</div></td><td class="ms-wiki-columnSpacing" style="width&#58;33.3%;"><div class="ms-rte-layoutszone-outer" style="width&#58;100%;"><div class="ms-rte-layoutszone-inner"></div>&#160;</div></td></tr><tr style="vertical-align&#58;top;"><td colspan="3"><div class="ms-rte-layoutszone-outer" style="width&#58;100%;"><div class="ms-rte-layoutszone-inner"></div>&#160;</div></td></tr></tbody></table><span id="layoutsData" style="display&#58;none;">true,true,3</span></div>
```
The value in "WikiField" is a \<div> tag containing a \<table> tag (#layoutsTable) and a \<span> tag (#layoutsData).

The #layoutsTable descridbes the page layout in details and the #layoutsData declares whether the page has header and footer and how many columns it contains. If you are guided to this page by the tool, please make sure

* #layoutsTable is valid HTML and the number of tr and td tags match a supported layout in following list.
* The inner HTML of #layoutsData should be in format as \<whether the page has header>, \<whether the page has footer>, \<the number of columns in page body> .
* The layout described in #layoutsTable should match #layoutsData.

Once you identified the problem in above steps, you can fix it by modifying the value in "WikiField".

**!!Make sure you understand what you are doing and back up the page before you begin the fix steps. You may lose part of page content when you change the layout data. !!**

Sample using PnP powershell cmdlet
```powershell
    Set-PnPListItem -List sitepages -Identity 11111 -Values @{"WikiField"='<div class="ExternalClass0ABFFA5383FD485EA2B0F7885197795D"><table id="layoutsTable" style="width&#58;100%;"><tbody><tr style="vertical-align&#58;top;"><td colspan="3"><div class="ms-rte-layoutszone-outer" style="width&#58;100%;"><div class="ms-rte-layoutszone-inner"></div>&#160;</div></td></tr><tr style="vertical-align&#58;top;"><td style="width&#58;33.3%;"><div class="ms-rte-layoutszone-outer" style="width&#58;100%;"><div class="ms-rte-layoutszone-inner"><p><br></p></div></div></td><td class="ms-wiki-columnSpacing" style="width&#58;33.3%;"><div class="ms-rte-layoutszone-outer" style="width&#58;100%;"><div class="ms-rte-layoutszone-inner"></div>&#160;</div></td><td class="ms-wiki-columnSpacing" style="width&#58;33.3%;"><div class="ms-rte-layoutszone-outer" style="width&#58;100%;"><div class="ms-rte-layoutszone-inner"></div>&#160;</div></td></tr><tr style="vertical-align&#58;top;"><td colspan="3"><div class="ms-rte-layoutszone-outer" style="width&#58;100%;"><div class="ms-rte-layoutszone-inner"></div>&#160;</div></td></tr></tbody></table><span id="layoutsData" style="display&#58;none;">true,true,3</span></div>'}
```

