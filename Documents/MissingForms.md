# Missing New/Edit/Display Forms for A Library/List

## Summary
This feature helps user restore new/edit/display forms if they are missing in the library/list

## Symptom
User cannot open 'Version history' for a file, and it shows an error as below. This generally happens if the document library is migrated from on-prem to online. 
![image](https://user-images.githubusercontent.com/79626459/185387125-1282f982-b7c8-4fd2-948f-6e896b403f75.png)

## Example

* Enter the affected site url. Select the library/list to check. Click 'Check Issues' button.
![image](https://user-images.githubusercontent.com/79626459/185387601-f140bb81-cc86-4f72-9bcc-133e49b045d4.png)

* It shows in red if any form is missing. Click 'Fix Issues' to recreate the form. 
![image](https://user-images.githubusercontent.com/79626459/185387814-fd927508-6aaa-461e-ba5f-091697983fb3.png)

* Check issues again, and the form is now restored. In the library, the version history can also be displayed. 
![image](https://user-images.githubusercontent.com/79626459/185388212-eb19456e-d34b-4a94-beb3-d91f79eaf9e9.png)
![image](https://user-images.githubusercontent.com/79626459/185388402-ac9e8d21-a5b0-42df-809d-1d12bc7f6209.png)

## More Information

The feature requires 'custom script' to be allowed. Reference: https://docs.microsoft.com/en-us/sharepoint/allow-or-prevent-custom-script#to-allow-custom-script-on-other-sharepoint-sites
