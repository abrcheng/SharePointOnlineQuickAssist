import * as React from 'react';
import {  
    PrimaryButton    
  } from 'office-ui-fabric-react/lib/index';
export default class UserProfileEmailQA extends React.Component
{
    public render():React.ReactElement<{}>
    {
        return (
            <div>
                  <PrimaryButton
                      text="Check UserProfile Email"
                      style={{ display: 'block', marginTop: '10px' }}
                      onClick={() => {alert("clicked"); }}
                    />
            </div>
        );
    }
}