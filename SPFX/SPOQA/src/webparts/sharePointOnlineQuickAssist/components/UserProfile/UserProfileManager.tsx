import * as React from 'react';
import * as strings from 'SharePointOnlineQuickAssistWebPartStrings';
import {  
    PrimaryButton    
  } from 'office-ui-fabric-react/lib/index';
import { ISharePointOnlineQuickAssistProps } from '../ISharePointOnlineQuickAssistProps';
export default class UserProfileManagerQA extends React.Component<ISharePointOnlineQuickAssistProps>
{
    public render():React.ReactElement<ISharePointOnlineQuickAssistProps>
    {
        return (
            <div>
                  <PrimaryButton
                      text={strings.UPM_CheckManager}
                      style={{ display: 'block', marginTop: '10px' }}
                      onClick={() => {alert("clicked"); }}
                    />
            </div>
        );
    }
}