import * as React from 'react';
import {  
    PrimaryButton    
  } from 'office-ui-fabric-react/lib/index';

export default class SearchLibraryQA extends React.Component
{
    public render():React.ReactElement<{}>
    {
        return (
            <div>
                  <PrimaryButton
                      text="Check Search Library"
                      style={{ display: 'block', marginTop: '10px' }}
                      onClick={() => {alert("clicked"); }}
                    />
            </div>
        );
    }
}