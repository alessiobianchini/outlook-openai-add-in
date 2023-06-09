/* eslint-disable no-undef */
import { Label, MessageBar, MessageBarType, PrimaryButton, Spinner, SpinnerSize, TextField } from "@fluentui/react";
import * as React from "react";

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
        const { isOfficeInitialized } = this.props;

        let bodyBackgroundColor = 'inherit';
        let bodyForegroundColor = 'inherit';
        let controlBackgroundColor = 'inherit';
        let controlForegroundColor = 'inherit';

        if (Office.context && Office.context.officeTheme) {
            bodyBackgroundColor = Office.context.officeTheme.bodyBackgroundColor;
            bodyForegroundColor = Office.context.officeTheme.bodyForegroundColor;
            controlBackgroundColor = Office.context.officeTheme.controlBackgroundColor;
            controlForegroundColor = Office.context.officeTheme.controlForegroundColor;
            if (document.getElementsByTagName('body')) {
                document.getElementsByTagName('body')[0].style['background-color'] = bodyBackgroundColor;
                document.getElementsByTagName('body')[0].style['color'] = bodyForegroundColor;
            }
        }

        const mainStyle = {
            display: 'flex',
            'flex-direction': 'column',
            'align-items': 'center'
        }

        if (!isOfficeInitialized) {
            return (<Spinner size={SpinnerSize.large} label={'Loading...'} />
            );
        }

        return (
            <main style={mainStyle}>
                <p >
                    <Label style={{ color: 'inherit', backgroundColor: 'inherit' }}>OpenAI API token configuration</Label>
                </p>
                <TextField
                    style={{ color: 'inherit', backgroundColor: 'inherit !important' }}
                    value={(this.state as any).token}
                    onChange={this.handleChange}
                    inputClassName="text-field"
                />
                <p>
                    <PrimaryButton
                        className='btn'
                        onClick={this.saveSettings}
                    >
                        Save API token
                    </PrimaryButton>
                </p>
                <p>
                    {!!((this.state as any).saved) && <SuccessExample title={(this.state as any).saved} />}
                    {!!((this.state as any).error) && <ErrorExample title={(this.state as any).error} />}
                </p>
                <p >
                    <Label style={{ color: 'inherit', backgroundColor: 'inherit' }}>
                        <a href="https://platform.openai.com/account/api-keys" target="_blank" rel="noopener noreferrer">
                            Get your OpenAI API key
                        </a>
                    </Label>
                </p>
            </main>
        );
    }
}