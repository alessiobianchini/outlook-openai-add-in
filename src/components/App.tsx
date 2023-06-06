/* eslint-disable no-undef */
import { Button, Input } from "@fluentui/react-components";
import * as React from "react";
import Progress from "./Progress";

/* global require */

export interface AppProps {
    title: string;
    isOfficeInitialized: boolean;
}

export interface AppState {
}

export default class App extends React.Component<AppProps, AppState> {

    constructor(props) {
        super(props);

        this.state = {
            token: Office.context.roamingSettings.get('openApiToken')
        };

        this.saveSettings = this.saveSettings.bind(this);
    }

    saveSettings() {
        Office.context.roamingSettings.set('openApiToken', (this.state as any).token);
        this.save();
        var itoken = Office.context.roamingSettings.get('openApiToken');
        var tt = itoken;
    }

    save() {
        Office.context.roamingSettings.saveAsync(function (result) {
            if (result.status !== Office.AsyncResultStatus.Succeeded) {
                console.error(`Action failed with message ${result.error.message}`);
            } else {
                console.log(`Settings saved with status: ${result.status}`);
            }
        });
    }

    handleChange = e => {
        this.setState({token: e.target.value})
    };

    render() {
        const { isOfficeInitialized, title } = this.props;

        if (!isOfficeInitialized) {
            return (
                <Progress
                    title={title}
                    logo={require("./../../assets/logo-filled.png")}
                    message="Please sideload your addin to see app body."
                />
            );
        }

        return (
            <div className="ms-welcome">
                <main className="ms-welcome__main">
                    <h2 className="ms-font-xl ms-fontWeight-semilight ms-fontColor-neutralPrimary ms-u-slideUpIn20">
                        AI Assistant
                    </h2>

                    <p className="ms-font-l ms-fontWeight-semilight ms-fontColor-neutralPrimary ms-u-slideUpIn20">
                        OpenAI token configuration
                    </p>
                    <p>
                        <Input value={(this.state as any).token} onChange={this.handleChange} />
                    </p>
                    <p>
                        <Button
                            appearance="primary"
                            onClick={this.saveSettings}
                        >
                            Save token
                        </Button>
                    </p>
                </main>
            </div>
        );
    }
}