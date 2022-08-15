# OneDrive Lock Icon
## Summary
The OneDrive lock icon issue can be caused by a lot of settings, sometimes the support engineer and tenant admin may forget to check some settings, so we developed this feature for helping support engineer and tenant admin to check related settings quickly.
![image](https://user-images.githubusercontent.com/21354416/184605424-eeec9045-3c92-47d9-a48b-59a911635289.png)
![image](https://user-images.githubusercontent.com/21354416/184605204-1e59830a-3148-4a54-9f52-64a49c1925ed.png)

## Example
* Select the "OneDrive lock icon", fill the affected user and affected site, select the affected library,
![image](https://user-images.githubusercontent.com/21354416/184606521-9fcc4569-1e0b-4748-a8a6-5593e116c8df.png)
* Click "Check Issues" button and wait for checking complete,
![image](https://user-images.githubusercontent.com/21354416/184607017-b725515a-4400-488a-afc9-44f1fb0c6a67.png)
* Check detected issues (in red) and click "Show Remedy Steps" button for checking remedy steps,
![image](https://user-images.githubusercontent.com/21354416/184607370-46f38c99-da26-4301-b591-c793df684a3f.png)
* Open the link in the remedy step in new tab and fix the settings (e.g. in this demo, need to remove the validation message and turn on the "Offline Client Availability") accordingly,
![image](https://user-images.githubusercontent.com/21354416/184607690-06191f4b-4fcc-402b-abc5-3f3ed8e25ab2.png)
![image](https://user-images.githubusercontent.com/21354416/184607787-1fbc96e8-6a5c-4e6a-a2e0-1fc02fb04259.png)

## More Information (this feature will check below settings)
* Offline Client Availability for the library has been set to true.
* Require Check Out for the library has been set to false.
* Draft Item Security of this library has been set to Any user who can read items.
* Content Approval of this library has been set to false.
* Validation formula/message of this library is null/null.
* Validation formula/message of all columns are null/null.
* The "Offline Client Availability" of the site(including parent sites) has been set to false.
* Limited-access user permission lockdown mode of the site collection hasn't been enabled.
* The affected user has enough permission to edit the library.
