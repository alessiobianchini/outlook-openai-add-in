/* eslint-disable no-undef */
import { Label, MessageBar, MessageBarType, PrimaryButton, TextField } from "@fluentui/react";
import * as React from "react";
import Progress from "./Progress";

/* global require */

export interface AppProps {
    title: string;
    isOfficeInitialized: boolean;
}


export interface MessageProps {
    title: string;
}

export interface AppState {
}

const SuccessExample = (props: MessageProps) => (
    <MessageBar
        messageBarType={MessageBarType.success}
        isMultiline={false}
        dismissButtonAriaLabel="Close"
    >
        {props.title}
    </MessageBar>
);


const ErrorExample = (props: MessageProps) => (
    <MessageBar
        messageBarType={MessageBarType.error}
        isMultiline={false}
        dismissButtonAriaLabel="Close"
    >
        {props.title}
    </MessageBar>
);

export default class App extends React.Component<AppProps, AppState> {

    constructor(props) {
        super(props);

        this.state = {
            token: Office.context.roamingSettings.get('openApiToken'),
            saved: null,
            error: null
        };

        this.saveSettings = this.saveSettings.bind(this);
    }

    saveSettings() {
        Office.context.roamingSettings.set('openApiToken', (this.state as any).token);
        Office.context.roamingSettings.saveAsync((result) => {
            if (result.status !== Office.AsyncResultStatus.Succeeded) {
                this.setState({ saved: null, error: `Save failed with message ${result.error.message}` });
                console.error(`Action failed with message ${result.error.message}`);
            } else {
                console.log(`Settings saved with status: ${result.status}`);
                this.setState({ saved: 'Open AI token saved!', error: null });
            }
        });
    }

    handleChange = e => {
        this.setState({ token: e.target.value })
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
                    <Label>
                        <h2>
                            AI Assistant
                        </h2>
                    </Label>

                    <p >
                        <Label>OpenAI token configuration</Label>
                    </p>
                    <TextField value={(this.state as any).token} onChange={this.handleChange} />
                    <p>
                        <PrimaryButton
                            className='btn'
                            onClick={this.saveSettings}
                        >
                            Save token
                        </PrimaryButton>
                    </p>
                    <p>
                        {!!((this.state as any).saved) && <SuccessExample title={(this.state as any).saved} />}
                        {!!((this.state as any).error) && <ErrorExample title={(this.state as any).error} />}
                    </p>
                    <p >
                        <a href="https://platform.openai.com/account/api-keys" target="_blank" rel="noopener noreferrer">
                            Get your OpenAI API key
                        </a>
                    </p>
                </main>
            </div>
        );
    }
}