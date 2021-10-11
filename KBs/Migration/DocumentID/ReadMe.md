**Issue Description,**
The document ID may can't be generated in time after migrating a lot of documents to SharePoint Online (SPO), the SPO DocID assignment job is a async timer job process and runs for a set duration in the day. It is shared across all customers in that farm. So there is no guarantee on how many day/months it will take if there are millions of documents and there is no quick backend way to assign the Ids. 
![image](https://user-images.githubusercontent.com/21354416/136776989-c67b11af-c811-4c99-abda-d33f1fe7c8fc.png)


**Solution,**
Run the script TouchDocumentIds.ps1 for trigger the document ID assignment, please note this script will trigger the workflow or PowerAutomate flow which associated with the library,
![image](https://user-images.githubusercontent.com/21354416/136777288-e358cfb0-ce05-4ed6-ac76-c4b2a11e8bc3.png)
![image](https://user-images.githubusercontent.com/21354416/136777331-99c63615-43cf-4dd3-9a49-537dd239eaff.png)

And you can run it in parallel for different libraries if necessary.
Then run the script GetDocIDReport.ps1 for getting document ID assignment report,

![image](https://user-images.githubusercontent.com/21354416/136783341-44fd2dcb-72b8-4be0-9b74-04e28cb462d5.png)
![image](https://user-images.githubusercontent.com/21354416/136777527-9272041f-e728-4a9c-95a2-7455195eeb72.png)
![image](https://user-images.githubusercontent.com/21354416/136777560-ec1ff9fa-8266-4b61-a707-50b50adb7ec0.png)

**Notes,**
1. These two scripts are depended on SharePoint Online PnP PowerShell https://docs.microsoft.com/en-us/powershell/sharepoint/sharepoint-pnp/sharepoint-pnp-cmdlets
2. These two scripts are provided "AS IS"
