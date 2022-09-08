import * as React from 'react';
import { IExoQuickAssistProps } from '../IExoQuickAssistProps';
import styles from '../ExoQuickAssist.module.scss';
import {  
    PrimaryButton,    
    TextField,
    Label
  } from 'office-ui-fabric-react/lib/index';

import * as strings from 'ExoQuickAssistWebPartStrings';

export default class TestTeamsQA extends React.Component<IExoQuickAssistProps>
{
    public state = {         
        affectedUser:this.props.currentUser.loginName,
        isNeedFix:false       
      };

    public render(): React.ReactElement<IExoQuickAssistProps> {
        return (
            <div>
              <div className={ styles.row }>
                <div className={ styles.column }>                 
                  <TextField                          
                          label={strings.AffectedUser}
                          multiline={false}
                          onChange={(e)=>{let text:any = e.target; this.setState({affectedUser:text.value});}}
                          value={this.state.affectedUser}
                          required={true}                                                
                    />   
                    <Label>e.g. John@contoso.com </Label>                  
                    
                    <PrimaryButton
                        text={strings.CheckIssue}
                        style={{ display: 'inline', marginTop: '10px' }}
                        onClick={() => {this.Check();}}
                      />
                      
                      { this.state.isNeedFix ? 
                      <PrimaryButton
                        text={strings.ShowRemedySteps}
                        style={{ display: 'inline', marginTop: '10px', marginLeft:"20px"}}             
                        onClick={() => {this.ShowRemedy();}}
                      />: null}
                  </div>
              </div>
          </div>
          );
    }

    private Check()
    {

    }

    private ShowRemedy()
    {

    }
}