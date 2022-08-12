# Restore Items
## Summary
A lot of users complain about they can't filter and restore items from recycle bin in SharePoint Online, so we developed this feature which allow user to filter/restore/export items from SharePoint Online recycle bin.

## Example
* Filter the items from recycle bin by delete date, delete by, path,

![image](https://user-images.githubusercontent.com/21354416/155688589-199fc965-1333-4073-82b6-677444497a36.png)

* After filtering items, user can restore all filtered items by clicking the restore button,
![image](https://user-images.githubusercontent.com/21354416/155689019-f91ba251-aace-4671-8c6a-60d489debc87.png)

* After filtering items can export filtered items by clicking export button,
![image](https://user-images.githubusercontent.com/21354416/155689228-e8bc0d3b-1cf6-4b48-904a-c23cfcfa3e83.png)

* The skip option for existing items when restoring has been added into "Restore Items",
  ![image](https://user-images.githubusercontent.com/21354416/184311082-8e61bc09-ca96-46d2-92db-f4ff4d5ea0b6.png)
 ![image](https://user-images.githubusercontent.com/21354416/184311738-a63cb573-f6c6-4fc7-ad98-25e1527f7631.png)

*	By default this option is off, it will affect the performance if there are some large libraries need to be scanned.
*	If this option is off, then existing column will be filled as false by default, but please remember that it is just mean the tool haven’t checked it.
*	When the option is selected, existing items will be skipped when restoring,
![image](https://user-images.githubusercontent.com/21354416/184311656-e4649ac6-4e4c-457b-bc0e-fb39ae5a2efd.png)


*	If some items are existing but the option is not enable then it may cause error message when restoring,
![image](https://user-images.githubusercontent.com/21354416/184311583-e26fa5c0-4254-4849-8af2-e203105033c4.png)

## More Information
* Queried xxx items means there are totally xxx items in the first and second recycle bins of the site
* Filtered xxx items means there are xxx matched these filters
* Skipped xxx items (existing) means there are xxx items already existing, they are skipped when restoring
