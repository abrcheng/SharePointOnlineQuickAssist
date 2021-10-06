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
        affectedSite:this.props.webAbsoluteUrl,
        affectedUser:this.props.currentUser.email,
        isWebThere:false,
        isWebNoIndex:false,
        isChecked:false
    };

    public render():React.ReactElement<ISharePointOnlineQuickAssistProps>
    {        
        return (
            <div>
                  <TextField
                        label="Affected Site:"
                        multiline={false}
                        onChange={(e)=>{let text:any = e.target; this.setState({affectedSite:text.value}); this.setState({isChecked:false});}}
                        value={this.state.affectedSite}
                        required={true}                                                
                  />
                  <TextField
                        label="Affected User:"
                        multiline={false}
                        onChange={(e)=>{let text:any = e.target; this.setState({affectedUser:text.value});}}
                        value={this.state.affectedUser}
                        required={true}                                                
                  /> 
                    {this.state.affectedSite!="" && this.state.isChecked? 
                        <div id="SearchSiteResultSection">
                            <Label>Diagnose result:</Label>
                            {this.state.isWebThere?<Label style={{"color":"Green",marginLeft:20}}>Found site with URL {this.state.affectedSite}</Label>:
                                <Label style={{"color":"Red",marginLeft:20}} >Site with URL {this.state.affectedSite} doesn't exist</Label>}
                            {this.state.isWebThere?
                            <div>
                            {this.state.isWebNoIndex?<Label style={{"color":"Red",marginLeft:20}}>The nocrawl has been enabled for site {this.state.affectedSite}</Label>:
                                <Label style={{"color":"Green",marginLeft:20}}>Site {this.state.affectedSite} is searchable</Label>}
                            </div>:null}
                        </div>:null
                    }
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
      };
      console.log(sumObj);

      return sumObj;

    }

    private async getUserIDByEmail(email:string,siteUrl:string):Promise<number>
    {
        var url = `${siteUrl}/_api/web/siteusers?$filter=Email eq '${email}'`;
        var userData:any = await RestAPIHelper.GetQueryUser(siteUrl,this.props.spHttpClient);
        return userData.value[0].Id;
    }

    public async CheckSiteSearchSettings()
    {
        this.setState({isChecked:false});
        try
        {
          var siteSearch = await RestAPIHelper.GetSerchResults(this.props.spHttpClient, this.props.rootUrl, this.state.affectedSite, "Site");
          console.log(siteSearch);

          var sum = await this.GetJsonResults(siteSearch);

          if(sum.total == 0)
          {
            if(sum.totalNoDup == 0)
            {
              console.log(`No Serach Result for the site`);
              SPOQAHelper.ShowMessageBar("Error", "No Serach Result for the site."); 

              //Site is not searchable. Proceed to check more
              //1. Check if the site exists
              //2. Check if the site is crawled/indexed
              //3. Check if the user has permissions
              //4. Try searching the site with a different keyword

              //Check if the site exists
              try{
                var webInfo = await RestAPIHelper.GetWeb(this.props.spHttpClient, this.state.affectedSite);
                this.setState({isWebThere:webInfo});
              }
              catch(err)
              {
                SPOQAHelper.ShowMessageBar("Error",`Get exception when try to get web with error message ${err}`);
                return;
              }
              if(webInfo)
              {
                //Check if the site is crawled/indexed
                try
                {
                  var noCrawl = await RestAPIHelper.IsWebNoCrawl(this.props.spHttpClient, this.state.affectedSite);
                  this.setState({isWebNoIndex:noCrawl});
                }
                catch(err)
                {
                  SPOQAHelper.ShowMessageBar("Error",`Get exception when try to check IsWebNoCrawl with error message ${err}`);
                  return;
                }

                //Check if the user has permissions
                try
                {
                  //Get User Login Id by Email
                  //var userLoginId = this.getUserIDByEmail(this.state.affectedUser, this.state.affectedSite);

                  var userInfoSite = await RestAPIHelper.GetUserFromUserInfoList(this.state.affectedUser, this.props.spHttpClient, this.state.affectedSite);
                  //↑ Get error if user email contains "'"
                  console.log(`User info is ${userInfoSite.Email}`);
                }
                catch(err)
                {
                  SPOQAHelper.ShowMessageBar("Error",`Get exception when try to get user info with error message ${err}`);
                  return;
                }
              }
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

        this.setState({isChecked:true});
    }
}
