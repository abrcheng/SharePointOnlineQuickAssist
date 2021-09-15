import * as React from 'react';
import {  
    PrimaryButton    
  } from 'office-ui-fabric-react/lib/index';
export default class UserProfileTitleQA extends React.Component
{
    public render():React.ReactElement<{}>
    {
        return (
            <div>
                  <PrimaryButton
                      text="Check UserProfile Title"
                      style={{ display: 'block', marginTop: '10px' }}
                      onClick={() => {alert("clicked"); }}
                    />
            </div>
        );
    }
}