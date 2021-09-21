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
    public inputs = {
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
                        value={this.inputs.affectedSite}
                        required={true}                                                
                  />  
                  <PrimaryButton
                      text="Check Site"
                      style={{ display: 'inline', marginTop: '10px' }}
                      onClick={() => {this.CheckSiteSearchSettings();}}
                    />
            </div>
        );
    }

    public async CheckSiteSearchSettings()
    {
        try
        {
          var userInfoSite = await RestAPIHelper.GetSiteSearchResult(this.props.spHttpClient, this.inputs.affectedSite);
          console.log(`${userInfoSite}`);
        }
        catch(err)
        {
           this.forceUpdate();
           console.log(`Error`);
        }

        
    }
}
