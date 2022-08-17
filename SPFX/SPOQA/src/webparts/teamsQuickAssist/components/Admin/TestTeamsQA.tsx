import * as React from 'react';
import { ITeamsQuickAssistProps } from '../ITeamsQuickAssistProps';
import styles from '../TeamsQuickAssist.module.scss';
import {  
    PrimaryButton,    
    TextField,
    Label
  } from 'office-ui-fabric-react/lib/index';

export default class TestTeamsQA extends React.Component<ITeamsQuickAssistProps>
{
    public state = {         
        affectedUser:this.props.currentUser.loginName,
        isNeedFix:false       
      };

    public render(): React.ReactElement<ITeamsQuickAssistProps> {
        return (
            <div>
              <div className={ styles.row }>
                <div className={ styles.column }>                 
                  <TextField                          
                          label="Affected User:"
                          multiline={false}
                          onChange={(e)=>{let text:any = e.target; this.setState({affectedUser:text.value});}}
                          value={this.state.affectedUser}
                          required={true}                                                
                    />   
                    <Label>e.g. John@contoso.com </Label>                  
                    
                    <PrimaryButton
                        text="Check Issues"
                        style={{ display: 'inline', marginTop: '10px' }}
                        onClick={() => {this.Check();}}
                      />
                      
                      { this.state.isNeedFix ? 
                      <PrimaryButton
                        text="Show Remedy Steps"
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