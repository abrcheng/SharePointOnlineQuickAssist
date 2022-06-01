import {SPHttpClient,ISPHttpClientOptions} from '@microsoft/sp-http';
import RestAPIHelper from './RestAPIHelper';
import {Text } from '@microsoft/sp-core-library';
// Microsoft.SharePoint.Client.Search.Administration.DocumentCrawlLog
export default class SearchHelper
{
    public static async GetCrawlLogByRest(spHttpClient:SPHttpClient,webAbsoluteUrl:string, urlFilter:string)
    {
        let apiUrl = webAbsoluteUrl + "/_vti_bin/client.svc/ProcessQuery"; 
        let endDate =  new Date();
        //let startDate = new Date();
        //startDate.setDate(endDate.getDate() -30);
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
                                <Parameter Type="DateTime">0001-01-01T00:00:00.0000000</Parameter>
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

    public static async GetManagedProperties(spHttpClient:SPHttpClient,webAbsoluteUrl:string, workId:string)
    {
        // Get https://chengc.sharepoint.com/_api/search/query?querytext='WorkId%3a%22xxxxxx%22'&rowlimit=1&refiners='ManagedProperties(filter%3d5000%2f0%2f*)'&clienttype='ContentSearchRegular'
        // let apiUrl = `${webAbsoluteUrl}/_api/search/query?querytext='WorkId="${workId}"'&rowlimit=1&refiners='ManagedProperties(filter=5000/0/*)'`;
        // let apiUrl = `https://chengc.sharepoint.com/_api/search/query?querytext='Path%3ahttps%3a%2f%2fchengc.sharepoint.com%2fExcelLib%2fexcel2.xlsx'&rowlimit=1&refiners='ManagedProperties(filter%3d5000%2f0%2f*)'&sortlist='%5bdocid%5d%3aascending'&hiddenconstraints='WorkId%3a%22157214370457891573%22'&clienttype='ContentSearchRegular'`;
        let apiUrl = `${webAbsoluteUrl}/_api/search/postquery`;     
        const body: string = JSON.stringify({
            request: {
                ClientType: "ModernWebPart",                
                Querytext: `WorkId="${workId}"`,               
                RowLimit: 1,
                Refiners:'ManagedProperties(filter=5000/0/*)'          
            }});
          const headers: Headers = new Headers();
          headers.append('Accept', 'application/json;odata=nometadata');
          headers.append('Content-type', 'application/json;charset=utf-8');
          headers.append('OData-Version', '3.0');

        const httpClientOptions: ISPHttpClientOptions = {           
          body: body,
          headers:headers
        };

        var res:any = await spHttpClient.post(apiUrl, SPHttpClient.configurations.v1,httpClientOptions);
        var propertiesFirstRequest =[];
        if(res.ok)
        {
            res = await res.json();
            if(res.PrimaryQueryResult.RefinementResults.Refiners 
                && res.PrimaryQueryResult.RefinementResults.Refiners.length >0 
                && res.PrimaryQueryResult.RefinementResults.Refiners[0].Entries.length >0
                )
                {
                    var selectProperties = [];
                    var excludeProperties=["ClassificationContext","ClassificationLastScan","Color","ContentDatabaseId"];
                    res.PrimaryQueryResult.RelevantResults.Table.Rows[0].Cells.forEach(e=>{propertiesFirstRequest.push(e);});
                    res.PrimaryQueryResult.RefinementResults.Refiners[0].Entries.forEach(e=>{
                        var filterMps = propertiesFirstRequest.filter(p=>p.Key == e.RefinementName);
                        if(filterMps.length== 0 && excludeProperties.indexOf(e.RefinementName)<0)
                        {
                            selectProperties.push(e.RefinementName);
                        }
                    });
                    const queryAllMPBody: string = JSON.stringify({
                        request: {
                            ClientType: "ModernWebPart",                
                            Querytext: `WorkId="${workId}"`,               
                            RowLimit: 1,
                            SelectProperties:selectProperties          
                        }});

                    const queryAllMPhttpClientOptions: ISPHttpClientOptions = {           
                            body: queryAllMPBody,
                            headers:headers
                          };
                    
                    try
                    {
                        var allMPRes:any = await spHttpClient.post(apiUrl, SPHttpClient.configurations.v1,queryAllMPhttpClientOptions);                       
                        if(allMPRes.ok)
                        {
                            res = await allMPRes.json();
                        }
                        else
                        {
                            console.error("Failed to call /_api/search/postquery when selecting all managed properties");
                        }
                    }
                    catch(ex)
                    {
                        console.error(ex);
                    }
                }
        }
        else
        {
            console.error("Failed to call /_api/search/postquery, ManagedProperties(filter=5000/0/*)");
        }
        
        // res.PrimaryQueryResult.RelevantResults.RowCount
        // res.PrimaryQueryResult.RelevantResults.Table.Rows[0].Cells
        // res.PrimaryQueryResult.RelevantResults.Table.Rows[0].Cells[0].Value
        // res.PrimaryQueryResult.RelevantResults.Table.Rows[0].Cells[0].Key
        propertiesFirstRequest.forEach(e=>{
            var filterProrpties = res.PrimaryQueryResult.RelevantResults.Table.Rows[0].Cells.filter(pro=>pro.Key == e.Key);
            if(filterProrpties.length ==0)
            {
                res.PrimaryQueryResult.RelevantResults.Table.Rows[0].Cells.push(e);
            }
        });

        return res;
    }

    public static async GetCrawledProperties(spHttpClient:SPHttpClient,webAbsoluteUrl:string, workId:string)
    {
        // Get https://chengc.sharepoint.com/_api/search/query?querytext='WorkId%3a%22xxxxx%22'&rowlimit=1&refiners='CrawledProperties(filter%3d5000%2f0%2f*)'&clienttype='ContentSearchRegular'
        // let apiUrl = `${webAbsoluteUrl}/_api/search/query?querytext='WorkId="${workId}"'&rowlimit=1&refiners='CrawledProperties(filter=5000/0/*)'`;
        let apiUrl = `${webAbsoluteUrl}/_api/search/postquery`;   
        const body: string = JSON.stringify({
            request: {
                ClientType: "ModernWebPart",                
                Querytext: `WorkId="${workId}"`,               
                RowLimit: 1,
                Refiners:'CrawledProperties(filter=5000/0/*)'          
            }});
          const headers: Headers = new Headers();
          headers.append('Accept', 'application/json;odata=nometadata');
          headers.append('Content-type', 'application/json;charset=utf-8');
          headers.append('OData-Version', '3.0');

        const httpClientOptions: ISPHttpClientOptions = {           
          body: body,
          headers:headers
        };

        var res = await spHttpClient.post(apiUrl, SPHttpClient.configurations.v1,httpClientOptions);
        if(res.ok)
        {
            res = await res.json();
        }
        else
        {
            console.error("Failed to call /_api/search/postquery, CrawledProperties(filter=5000/0/*)");
        }
        
        // res.PrimaryQueryResult.RefinementResults.Refiners[0].Entries[0].RefinementName
        return await res;        
    }

    public static async CallOtherDiagnosticsAPIS(spHttpClient:SPHttpClient,webAbsoluteUrl:string, workId:string)
    {
        // 	 /_api/search/query?querytext='workid:{0}'&properties='QueryIdentityDiagnostics:true'&property='EnableDynamicGroups:true'&TrimDuplicates=false	    
        //  /_api/search/query?querytext='workid:{0}'&properties='3SRouted:false,QueryIdentityDiagnostics:true'&property='EnableDynamicGroups:true'&TrimDuplicates=false
        var apiTemplates = ["{0}/_api/search/query?querytext='workid:{1}'&properties='QueryIdentityDiagnostics:true'&property='EnableDynamicGroups:true'&TrimDuplicates=false",
                            "{0}/_api/search/query?querytext='workid:{1}'&properties='3SRouted:false,QueryIdentityDiagnostics:true'&property='EnableDynamicGroups:true'&TrimDuplicates=false"];
        for(var index=0; index<apiTemplates.length; index++)
        {
            var apiUrl = Text.format(apiTemplates[index], webAbsoluteUrl, workId);
            await spHttpClient.get(apiUrl, SPHttpClient.configurations.v1);
        }
    }
}