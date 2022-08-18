# Search issue for a specific document

## Summary
This feature helps user fix the issue when a specific document does not appear in the search results

## Example

* Enter the affected site url. Click 'Check Issues' button.
![image](https://user-images.githubusercontent.com/79626459/185375892-fe9ec6cc-a8d5-4d7b-a335-bf79d9643fdc.png)

* Example 1: the site is searchable
![image](https://user-images.githubusercontent.com/79626459/185376703-5fa27117-f267-441d-aa2b-82b05e88e8dc.png)

* If the site can be searched, will show the managed properties and crawled properties. 
![image](https://user-images.githubusercontent.com/79626459/185382866-e932a93c-1b62-45ee-b9a8-5663269e1994.png)

* These properties can be filtered and exported,
![image](https://user-images.githubusercontent.com/79626459/185383097-d9cef95f-fa7f-40a6-9a02-627f880d0542.png)

* Example 2: the site cannot be searched. It detected that 'Search and Offline Availability' is set to 'No'. And the user is not in the 'Members' group. It's suggested to follow the remedy steps in new tab and fix the issue. 
![image](https://user-images.githubusercontent.com/79626459/185377239-370380a8-5254-4395-bb76-1f22a77fc11d.png)

* The crawl logs can be checked by clicking 'Show Crawl Logs'. 
![image](https://user-images.githubusercontent.com/79626459/185380691-5eb48fd0-d7b6-428b-a3b1-018aeb6b0f66.png)

## More Information

The feature diagnoses and fixes the issue as follows:

* The site's and its sub/parent sites Nocrawl are enabled
* The user has permissions to the site
* If the site is a group site, check if the user is in the 'Members' group
* Crawl log of the site (need to grant permssion according to https://docs.microsoft.com/en-us/sharepoint/set-crawl-log-permissions)
* All managed properties of the site if it can be searched
* All crawled properties of the site if it can be searched
