# Permission issue
## Summary
In daily work we see a lot of customers reported permission issue, so we summary some common scenarios which may cause the performance issue and added the permission check function.
![image](https://user-images.githubusercontent.com/21354416/184810706-55938c64-fd2f-4328-85c1-3eee517cb75f.png)

## Example
a. Specify the user and objects (affected site, library, [full document URL](https://github.com/abrcheng/SharePointOnlineQuickAssist/tree/main/Documents/How%20to%20collect%20display%20url%20for%20files%20and%20list%20items)) which need to be checked,
![image](https://user-images.githubusercontent.com/21354416/184810516-572b0192-5615-49eb-a0f6-f1e8e7df77b9.png)

b. Click "Check Issues" button and wait for checking complete
![image](https://user-images.githubusercontent.com/21354416/160529848-f00bb12f-932a-4bd8-8fe2-dbefc6739467.png)

c. Click "Show Remedy Steps" button,
![image](https://user-images.githubusercontent.com/21354416/184810982-1e3f8619-52e8-4958-b6b7-5c7600b29b48.png)


d.  Open the remedy link in a new tab,
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
