Demo for adding SPOQA to ODB root site for global tenants


	• First navigate to https://yourDomain-my.sharepoint.com/_layouts/15/ManageFeatures.aspx and activate the Site Pages feature as in screenshot below.
<img width="928" alt="image" src="https://user-images.githubusercontent.com/89838160/173485332-ce8e4bff-f47e-41e2-a6b5-9a97ee8416da.png">


	• After the above step you can go to https://yourDomain-my.sharepoint.com/_layouts/15/viewlsts.aspx?view=14 to have a look.

<img width="916" alt="image" src="https://user-images.githubusercontent.com/89838160/173484787-4813135b-c0fd-4c9a-8c5c-2e63e3493ce8.png">

 
 
	• Then run the below cmdlets to add AppCatalogto ODB root site ( https://yourDomain-my.sharepoint.com)
		Connect-SPOService -url https://m365x026020-admin.sharepoint.com
		Add-SPOSiteCollectionAppCatalog -Site https://m365x026020-my.sharepoint.com 
<img width="933" alt="image" src="https://user-images.githubusercontent.com/89838160/173484811-1be8bb74-f676-4b00-84c8-59e37cd393ef.png">


	• Go back to https://yourDomain-my.sharepoint.com/_layouts/15/viewlsts.aspx?view=14
	• Select Apps for SharePoint > click upload to upload the SPOQA app.

<img width="949" alt="image" src="https://user-images.githubusercontent.com/89838160/173484842-22795a8d-88b2-421b-aca4-f36f4334bba3.png">

<img width="941" alt="image" src="https://user-images.githubusercontent.com/89838160/173484897-249a37cb-5d7f-4a0b-8a24-62722d6f0b28.png">


	• Click Deployon the next screen that pops up.
	• Go back to https://yourDomain-my.sharepoint.com/_layouts/15/viewlsts.aspx?view=14

<img width="957" alt="image" src="https://user-images.githubusercontent.com/89838160/173484934-8937c391-39d5-42a5-8461-74af1042b2b8.png">
<img width="986" alt="image" src="https://user-images.githubusercontent.com/89838160/173484972-d0aabd96-df4b-48cd-a282-8b7996bc33f3.png">



	• Go back to https://yourDomain-my.sharepoint.com/_layouts/15/viewlsts.aspx?view=14
  
  <img width="996" alt="image" src="https://user-images.githubusercontent.com/89838160/173485001-cc6c88af-54da-4eba-a6fc-077a7e91d32f.png">

	• Create a page and add SPOQA webpart in the page.

<img width="1037" alt="image" src="https://user-images.githubusercontent.com/89838160/173485030-637f116c-b87f-4df7-8c08-6901c7eb7139.png">

<img width="986" alt="image" src="https://user-images.githubusercontent.com/89838160/173485078-fc63df80-04b0-493d-b321-9311c5efae59.png">

