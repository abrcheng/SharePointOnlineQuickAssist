import * as React from 'react';
import {  
    PrimaryButton    
  } from 'office-ui-fabric-react/lib/index';
import { ISharePointOnlineQuickAssistProps } from './ISharePointOnlineQuickAssistProps';
export default class SearchDocumentQA extends React.Component<ISharePointOnlineQuickAssistProps>
{
    public render():React.ReactElement<ISharePointOnlineQuickAssistProps>
    {
        return (
            <div>
                  <PrimaryButton
                      text="Check Search Document"
                      style={{ display: 'block', marginTop: '10px' }}
                      onClick={() => {alert("clicked"); }}
                    />
            </div>
        );
    }
}