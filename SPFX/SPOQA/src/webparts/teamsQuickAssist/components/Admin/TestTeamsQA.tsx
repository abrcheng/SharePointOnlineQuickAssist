import * as React from "react";
import { ITeamsQuickAssistProps } from "../ITeamsQuickAssistProps";
import styles from "../TeamsQuickAssist.module.scss";
import {
  PrimaryButton,
  TextField,
  Label,
} from "office-ui-fabric-react/lib/index";

import GraphAPIHelper from "../../../Helpers/GraphAPIHelper";
import SPOQAHelper from "../../../Helpers/SPOQAHelper";
import * as strings from "TeamsQuickAssistWebPartStrings";
import { Text } from "@microsoft/sp-core-library";

export default class TestTeamsQA extends React.Component<ITeamsQuickAssistProps> {
  public state = {
    affectedUser: this.props.currentUser.loginName,
    isNeedFix: false,
  };

  public render(): React.ReactElement<ITeamsQuickAssistProps> {
    return (
      <div>
        <div className={styles.row}>
          <div className={styles.column}>
            <TextField
              label="Affected User:"
              multiline={false}
              onChange={(e) => {
                let text: any = e.target;
                this.setState({ affectedUser: text.value });
              }}
              value={this.state.affectedUser}
              required={true}
            />
            <Label>e.g. John@contoso.com </Label>

            <PrimaryButton
              text="Check Issues"
              style={{ display: "inline", marginTop: "10px" }}
              onClick={() => {
                this.Check();
              }}
            />

            {this.state.isNeedFix ? (
              <PrimaryButton
                text="Show Remedy Steps"
                style={{
                  display: "inline",
                  marginTop: "10px",
                  marginLeft: "20px",
                }}
                onClick={() => {
                  this.ShowRemedy();
                }}
              />
            ) : null}
          </div>
        </div>
      </div>
    );
  }

  private async Check() {
    try {
      SPOQAHelper.ResetFormStaus();
      let userInfo: any = await GraphAPIHelper.GetUserInfo(
        this.state.affectedUser,
        this.props.msGraphClient
      );
      let roles: any = await GraphAPIHelper.GetUserRoles(
        userInfo.id,
        this.props.msGraphClient
      );
      if (roles.value.length < 50) {
        SPOQAHelper.ShowMessageBar(
          "Success",
          Text.format(strings.RO_CheckPass, roles.value.length)
        );
      } else {
        SPOQAHelper.ShowMessageBar("Error", strings.RO_CheckFail);
        this.setState({isNeedFix: true});
      }
    } catch (err) {
      SPOQAHelper.ShowMessageBar("Error", strings.RO_APIError);
      console.log(err);
    }
  }

  private ShowRemedy() {
    try {
      SPOQAHelper.ResetFormStaus();
      SPOQAHelper.ShowMessageBar("Info",strings.RO_Remedy);
    }
    catch (err){
      console.log(err);
    }
  }
}
