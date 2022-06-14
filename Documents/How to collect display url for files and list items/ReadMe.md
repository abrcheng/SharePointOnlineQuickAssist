## Collect display url for a file
The display url for a file in library could be easily collected by clicking the [Copy direct link] button on detail panel.
![getfilepath](https://user-images.githubusercontent.com/102142347/173496710-a20bc968-5303-4e1b-a088-5c933dc3cdda.gif)


## Collect display url for a list item
To collect display url for a list item, we should collect the direct link in the same way first, and then modified the {id}_.000 to dispform.aspx?ID={id}.
![getitempath](https://user-images.githubusercontent.com/102142347/173496745-b6eb8eab-4eac-4910-8f05-a7e7ff2d40c8.gif)

e.g.

If the direct link to the list item was

>https://contoso.sharepoint.com/sites/SPOQA/Lists/List1/2_.000

please change it to

>https://contoso.sharepoint.com/sites/SPOQA/Lists/List1/dispform.aspx?ID=2
