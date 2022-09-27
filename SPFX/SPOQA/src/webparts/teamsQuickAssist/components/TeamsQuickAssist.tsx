import * as React from "react";
import styles from "./TeamsQuickAssist.module.scss";
import { ITeamsQuickAssistProps } from "./ITeamsQuickAssistProps";
import { escape } from "@microsoft/sp-lodash-subset";
import {
  ComboBox,
  Fabric,
  IComboBox,
  IComboBoxOption,
  mergeStyles,
  SelectableOptionMenuItemType,
  Spinner,
  MessageBar,
  MessageBarType,
} from "office-ui-fabric-react/lib/index";

import * as strings from "TeamsQuickAssistWebPartStrings";
import TestTeamsQA from "./Admin/TestTeamsQA";
import SPOQAHelper from "../../Helpers/SPOQAHelper";

const wrapperClassName = mergeStyles({
  selectors: {
    "& > *": { marginBottom: "20px" },
    "& .ms-ComboBox": { maxWidth: "300px" },
    "& .ms-ComboBox-option": { marginLeft: "15px" },
  },
});

const INITIAL_OPTIONS: IComboBoxOption[] = [
  {
    key: "AccountIssue",
    text: strings.AccountIssue,
    itemType: SelectableOptionMenuItemType.Header,
  },
  { key: "LoginCookieError", text: strings.LoginCookieError },
];

export default class TeamsQuickAssist extends React.Component<
  ITeamsQuickAssistProps,
  {}
> {
  public state = {
    selectedKey: "",
  };

  public render(): React.ReactElement<ITeamsQuickAssistProps> {
    const teamsQADetail = () => {
      switch (this.state.selectedKey) {
        case "LoginCookieError":
          return (
            <TestTeamsQA
              msGraphClient={this.props.msGraphClient}
              currentUser={this.props.currentUser}
              ctx={this.props.ctx}
            />
          );
        default:
          return <div id="NoContentPlaceHolder" />;
      }
    };

    return (
      <div className={styles.teamsQuickAssist}>
        <Fabric className={wrapperClassName} id="SPOAQFabric">
          <div className={styles.container}>
            <div className={styles.row} id="TeamsQAContainer">
              <div className={styles.column}>
                <span className={styles.title}>{strings.WebPartTitle}</span>
              </div>
            </div>

            <div className={styles.row} id="TeamsQAuestionsContainer">
              <div className={styles.column}>
                <div>
                  <ComboBox
                    defaultSelectedKey="-1"
                    label={strings.SelectIssueTip}
                    allowFreeform
                    autoComplete="on"
                    options={INITIAL_OPTIONS}
                    required={true}
                    onChange={(
                      ev: React.FormEvent<IComboBox>,
                      option?: IComboBoxOption
                    ): void => {
                      this.setState({ selectedKey: option.key });
                    }}
                  />
                </div>
              </div>
            </div>

            <div className={styles.row} id="TeamsQADetailContainer">
              <div>{teamsQADetail()}</div>
            </div>
            <div className={styles.row} id="TeamsQAStatusContainer">
              <div className={styles.column}>
                <div>
                  <Spinner
                    id="SPOQASpinner"
                    label="Checking..."
                    ariaLive="assertive"
                    labelPosition="left"
                    style={{ display: "none" }}
                  />
                  <div
                    id="SPOQAErrorMessageBarContainer"
                    style={{ display: "none" }}
                  >
                    <MessageBar
                      id="SPOQAErrorMessageBar"
                      messageBarType={MessageBarType.error}
                      isMultiline={true}
                      onDismiss={() => {
                        SPOQAHelper.Hide("SPOQAErrorMessageBarContainer");
                      }}
                      dismissButtonAriaLabel="Close"
                    >
                      SPOQAErrorMessageBar
                    </MessageBar>
                  </div>
                  <div
                    id="SPOQASuccessMessageBarContainer"
                    style={{ display: "none" }}
                  >
                    <MessageBar
                      id="SPOQASuccessMessageBar"
                      messageBarType={MessageBarType.success}
                      isMultiline={true}
                      onDismiss={() => {
                        SPOQAHelper.Hide("SPOQASuccessMessageBarContainer");
                      }}
                      dismissButtonAriaLabel="Close"
                    >
                      SPOQASuccessMessageBar
                    </MessageBar>
                  </div>
                  <div
                    id="SPOQAWarningMessageBarContainer"
                    style={{ display: "none" }}
                  >
                    <MessageBar
                      id="SPOQAWarningMessageBar"
                      messageBarType={MessageBarType.warning}
                      isMultiline={true}
                      onDismiss={() => {
                        SPOQAHelper.Hide("SPOQAWarningMessageBarContainer");
                      }}
                      dismissButtonAriaLabel="Close"
                    >
                      SPOQAWarningMessageBar
                    </MessageBar>
                  </div>
                  <div
                    id="SPOQAInfoMessageBarContainer"
                    style={{ display: "none" }}
                  >
                    <MessageBar
                      id="SPOQAInfoMessageBar"
                      messageBarType={MessageBarType.info}
                      isMultiline={true}
                      onDismiss={() => {
                        SPOQAHelper.Hide("SPOQAInfoMessageBarContainer");
                      }}
                      dismissButtonAriaLabel="Close"
                    >
                      SPOQAInfoMessageBar
                    </MessageBar>
                  </div>
                </div>
              </div>
            </div>
          </div>
        </Fabric>
      </div>
    );
  }
}
