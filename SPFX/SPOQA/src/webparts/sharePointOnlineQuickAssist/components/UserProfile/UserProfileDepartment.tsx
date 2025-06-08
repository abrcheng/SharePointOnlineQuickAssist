import * as React from 'react';
import * as strings from 'SharePointOnlineQuickAssistWebPartStrings';
import {  
    DefaultButton    
  } from 'office-ui-fabric-react/lib/index';
import { ISharePointOnlineQuickAssistProps } from '../ISharePointOnlineQuickAssistProps';
export default class UserProfileDepartmentQA extends React.Component<ISharePointOnlineQuickAssistProps>
{
    public render():React.ReactElement<ISharePointOnlineQuickAssistProps>
    {
        return (
            <div>
                  <DefaultButton
                      text={strings.UPD_CheckDepartment}
                      style={{ display: 'block', marginTop: '10px' }}
                      onClick={() => {alert("clicked"); }}
                    />
            </div>
        );
    }
}