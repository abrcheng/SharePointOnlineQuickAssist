import * as React from 'react';
import * as strings from 'SharePointOnlineQuickAssistWebPartStrings';
import {  
    PrimaryButton    
  } from 'office-ui-fabric-react/lib/index';
import { ISharePointOnlineQuickAssistProps } from '../ISharePointOnlineQuickAssistProps';
export default class UserProfileEmailQA extends React.Component<ISharePointOnlineQuickAssistProps>
{
    public render():React.ReactElement<ISharePointOnlineQuickAssistProps>
    {
        return (
            <div>
                  <PrimaryButton
                      text={strings.UPE_CheckEmail}
                      style={{ display: 'block', marginTop: '10px' }}
                      onClick={() => {alert("clicked"); }}
                    />
            </div>
        );
    }
}