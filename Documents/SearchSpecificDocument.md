# Search issue for a specific document

## Summary
This feature helps user fix the issue when a specific document does not appear in the search results

## Example

Enter the affected site url and select the affected list/library. Then enter the full URL/path of the affected document. Click 'Check Search Document'.

<img src=../SPFX/SPOQA/asset/NoCrawl.JPG>


In this example, it detected that the library's nocrawl was enabled. Just click 'Fix issues' and it will be fixed automatically.

<img src=../SPFX/SPOQA/asset/FixedNoCrawl.JPG>


## More Information

The feature diagnoses and fixes the issue as follows:

* The site's Nocrawl is enabled
* The affected library/list's Nocrawl is enabled
* The DispForm.aspx of the affected library/list is missing
* The affected document is having no major version
