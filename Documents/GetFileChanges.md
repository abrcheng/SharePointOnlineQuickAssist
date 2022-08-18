# Get File Changes (New or Update)
## Summary
This feature helps get all or filtered file changes for a given site

## Example

* Enter the site url to get the changes. Click 'Get Files" and it will start to query.
![image](https://user-images.githubusercontent.com/79626459/185389396-bec89e6c-76e1-4271-b6e6-3c68e3568afb.png)

* It shows the changes list with the following format. 
![image](https://user-images.githubusercontent.com/79626459/185389807-38be33d1-4faf-4442-81ff-ea2048c1ce69.png)

* User can also filter the results by 'Modified User', 'Path', 'Start Date' or 'End Date".
![image](https://user-images.githubusercontent.com/79626459/185390292-bb9bf6f4-cd73-46fb-933e-9eb6ecbcab86.png)

* The results can be exported to a CSV file. 
![image](https://user-images.githubusercontent.com/79626459/185390615-f1f7ab00-1716-490b-8e0d-ac3f7b59dc64.png)

## More Information

The feature is based on 'delta query'. Reference: https://docs.microsoft.com/en-us/graph/delta-query-overview

It doesn't work for 21V tenants becasue 'delta query' is not supported there. 
