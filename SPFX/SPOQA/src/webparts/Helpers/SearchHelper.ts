import {SPHttpClient,ISPHttpClientOptions} from '@microsoft/sp-http';
// Microsoft.SharePoint.Client.Search.Administration.DocumentCrawlLog
export default class SearchHelper
{
    public static async GetCrawlLogByRest(spHttpClient:SPHttpClient,webAbsoluteUrl:string, urlFilter:string)
    {
        let apiUrl = webAbsoluteUrl + "/_vti_bin/client.svc/ProcessQuery"; 
        let endDate =  new Date();
        let startDate = new Date();
        startDate.setDate(endDate.getDate() -30);
        let userData =`        
                    <Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName=".NET Library" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009">
                    <Actions>
                        <ObjectPath Id="2" ObjectPathId="1" />
                        <ObjectPath Id="4" ObjectPathId="3" />
                        <ObjectPath Id="6" ObjectPathId="5" />
                        <Query Id="7" ObjectPathId="5">
                            <Query SelectAllProperties="true">
                                <Properties />
                            </Query>
                        </Query>
                        <Method Name="GetCrawledUrls" Id="8" ObjectPathId="5">
                            <Parameters>
                                <Parameter Type="Boolean">false</Parameter>
                                <Parameter Type="Int64">100</Parameter>
                                <Parameter Type="String">${urlFilter}</Parameter>
                                <Parameter Type="Boolean">true</Parameter>
                                <Parameter Type="Int32">-1</Parameter>
                                <Parameter Type="Int32">-1</Parameter>
                                <Parameter Type="Int32">-1</Parameter>
                                <Parameter Type="DateTime">${startDate.toISOString()}</Parameter>
                                <Parameter Type="DateTime">${endDate.toISOString()}</Parameter>
                            </Parameters>
                        </Method>
                    </Actions>
                    <ObjectPaths>
                        <StaticProperty Id="1" TypeId="{3747adcd-a3c3-41b9-bfab-4a64dd2f1e0a}" Name="Current" />
                        <Property Id="3" ParentId="1" Name="Site" />
                        <Constructor Id="5" TypeId="{5c5cfd42-0712-4c00-ae49-23b33ba34ecc}">
                            <Parameters>
                                <Parameter ObjectPathId="3" />
                            </Parameters>
                        </Constructor>
                    </ObjectPaths>
                    </Request>`;
      
        let spOpts = {  
            headers: {  
                'Accept': 'application/json;odata=nometadata',  
                'Content-type': 'application/json;odata=verbose',  
                'odata-version': '',  
            },  
            body: userData 
        };  
        var res = await spHttpClient.post(apiUrl, SPHttpClient.configurations.v1, spOpts); 
        if(res.ok)
        {
          var responseJson = await res.json();
          for(var index=0; index < responseJson.length; index++)
          {
              if(responseJson[index]._ObjectType_ == "SP.SimpleDataTable")
              {
                console.log(`GetCrawlLogByRest get rows ${responseJson[index].Rows.length} via API ${apiUrl} with url fitler ${urlFilter}`);        
                responseJson = responseJson[index];
              }
          }

          console.log(`GetCrawlLogByRest done for API url ${apiUrl}`);          
          return await responseJson;
        } 
    }
}