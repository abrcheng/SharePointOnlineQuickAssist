# Search issue for a specific document

## Summary
This feature helps user fix the issue when a specific document does not appear in the search results

## Example

* Enter the affected site url and select the affected list/library. Then enter the full URL/path of the affected document. Click 'Check Issues' button.
![image](https://user-images.githubusercontent.com/21354416/184611775-f1e9e4f7-d02e-44f2-a384-5cd495da783d.png)

* In this example, it detected that the library's nocrawl was enabled. Just click 'Show Remedy Steps' and it will show remedy steps,
![image](https://user-images.githubusercontent.com/21354416/184611984-bc0285ba-e794-4e53-8ef5-1b50a9ae2880.png)

* Open the link in the remedy step in new tab and fix the settings (e.g. in this demo, need to trun off the no crawl in list advanced settings ) accordingly,
![image](https://user-images.githubusercontent.com/21354416/184612372-d2109f4f-1379-4b12-82fc-912e77a9b0d7.png)

* And the crawl log can be checked by below ,
 
![image](https://user-images.githubusercontent.com/21354416/171319876-02339ed1-8015-4a8f-9043-da93c89d99da.png)

* If the document can be searched, will show all managed properties,
![image](https://user-images.githubusercontent.com/21354416/171320028-d9aab9f0-1f68-4841-b9e9-d4108ce75f46.png)

 * If the document can be searched, will show all crawled properties as well

* These properties can be filtered and exported,
![image](https://user-images.githubusercontent.com/21354416/171320213-d81bd049-485a-4eb2-a6e5-583eb15c29b4.png)
![image](https://user-images.githubusercontent.com/21354416/171320278-a718070a-7d26-441d-8e8c-e6dcc6044bcf.png)
![image](https://user-images.githubusercontent.com/21354416/171320453-162bd9ef-1912-4b2e-9e45-630f8014ea49.png)

## More Information

The feature diagnoses and fixes the issue as follows:

* The site's Nocrawl is enabled
* The affected library/list's Nocrawl is enabled
* The DispForm.aspx of the affected library/list is missing
* The affected document is having no major version
* Crawl log of the document (need to grant permssion according to https://docs.microsoft.com/en-us/sharepoint/set-crawl-log-permissions)
* All managed properties of the document if the doucment can be searched
* All crawled properties of the document if the doucment can be searched
