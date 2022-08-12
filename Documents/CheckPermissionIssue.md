# Permission issue
## Summary
In daily work we see a lot of customers reported permission issue, so we summary some common scenarios which may cause the performance issue and added the permission check function.

![image](https://user-images.githubusercontent.com/21354416/161952654-8706562b-0b6d-4eb7-a1ef-a9bce034af64.png)

## Example
a. Specify the user and object which need to be checked, click "check issue" button 
![image](https://user-images.githubusercontent.com/21354416/160529848-f00bb12f-932a-4bd8-8fe2-dbefc6739467.png)

b. Click "show remedy" button,
![image](https://user-images.githubusercontent.com/21354416/160530111-097ad641-db02-4817-bc11-1aaf80ebbc82.png)

c.  Open the remedy link,
![image](https://user-images.githubusercontent.com/21354416/160530199-18ec4d8d-d132-4263-b8b2-dd50b6960d92.png)

## More Information
The feature diagnoses and fixes the issue as follows:
1. There isn't any documents without check-in version in the library.
2. The file can be found.
3. The user has read permission on the document
4. The document is not in draft version
5. The library hasn't been set to only the author can read/write the item
6. Limited-access user permission lockdown mode of the site collection has been disabled
7. The affected user has view permission on the library.
8. The customization of the modern/classic page 
