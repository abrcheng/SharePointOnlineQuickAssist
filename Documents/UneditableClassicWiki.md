# Un-editable Classic Wiki

## Summary
This feature helps user to detect classic wiki page's layout issue which cause the whole page including the ribbon menu being non-interactive in edit mode.
[More Details](https://github.com/abrcheng/SharePointOnlineQuickAssist/blob/main/KBs/UneditableClassicWiki/ReadMe.md)

## Usage
Input the page URL and click [Check Issues].

## Example
The feature will detect the page layout by parsing HTML elements and compare the detected layout to page's layout declaration.
In the example below, the page layout was detected as "Three column with header and footer" while the layout declaration was not a valid value.
![image](https://user-images.githubusercontent.com/102142347/185521576-23b73935-b74f-4f2e-aa70-d2dc8e4041dc.png)
