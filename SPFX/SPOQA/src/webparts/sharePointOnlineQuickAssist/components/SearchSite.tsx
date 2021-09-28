import * as React from 'react';
import {  
    PrimaryButton,
    Text,
    TextField,
    Label
  } from 'office-ui-fabric-react/lib/index';
import GraphAPIHelper from '../../Helpers/GraphAPIHelper';  
import RestAPIHelper from '../../Helpers/RestAPIHelper';
import SPOQASpinner from '../../Helpers/SPOQASpinner';
import SPOQAHelper from '../../Helpers/SPOQAHelper';
import { ISharePointOnlineQuickAssistProps } from './ISharePointOnlineQuickAssistProps';
export default class SearchSiteQA extends React.Component<ISharePointOnlineQuickAssistProps>
{
    public state = {
        affectedSite:this.props.webAbsoluteUrl
    };

    public render():React.ReactElement<ISharePointOnlineQuickAssistProps>
    {        
        return (
            <div>
                  <TextField
                        label="Affected Site:"
                        multiline={false}
                        onChange={(e)=>{let text:any = e.target; this.setState({affectedSite:text.value});}}
                        value={this.state.affectedSite}
                        required={true}                                                
                  />  
                  <PrimaryButton
                      text="Check Site"
                      style={{ display: 'inline', marginTop: '10px' }}
                      onClick={() => {SPOQAHelper.ResetFormStaus();this.CheckSiteSearchSettings();}} //When click: Reset banner status & check if the site is searchable
                    />
            </div>
        );
    }

    private async GetJsonResults(JsonStr:string)
    {
      interface ResultsSum {
        queryModifcation: string;
        total: number;
        totalNoDup: number;
      }
      
      var sumObj:ResultsSum = {
        queryModifcation: "string",
        total: JsonStr['PrimaryQueryResult'].RelevantResults.TotalRows,
        totalNoDup: JsonStr['PrimaryQueryResult'].RelevantResults.TotalRowsIncludingDuplicates
      }
      console.log(sumObj);

      return sumObj;

    }

    public async CheckSiteSearchSettings()
    {
        try
        {
          var userInfoSite = await RestAPIHelper.GetSerchResults(this.props.spHttpClient, this.props.rootUrl, this.state.affectedSite, "Site");
          console.log(userInfoSite);

          var sum = await this.GetJsonResults(userInfoSite);

          if(sum.total == 0)
          {
            if(sum.totalNoDup == 0)
            {
              console.log(`No Serach Result for the site`);
              SPOQAHelper.ShowMessageBar("Error", "No Serach Result for the site."); 

              //Site is not searchable. Proceed to check more
              //1. Check if the site exists
              //2. Check if the user has permissions
              //3. Check if the site is crawled/indexed
              //4. Try searching the site with a different keyword

              



            }
            else
            {
              console.log(`Serach Result in Duplicate for the site`);
              SPOQAHelper.ShowMessageBar("Warning", "Serach Result in Duplicate for the site."); 
            }
          }
          else
          {
            console.log(`The site is searchable`);
            SPOQAHelper.ShowMessageBar("Success", "The site is searchable."); 
          }
        }
        catch(err)
        {
           this.forceUpdate();
           console.log(`Error`);
        }

        
    }
}
