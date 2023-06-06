import { initializeIcons } from "@fluentui/font-icons-mdl2";
import { ThemeProvider } from "@fluentui/react";
import * as React from "react";
import { createRoot } from 'react-dom/client';
import { AppContainer } from "react-hot-loader";
import App from "./components/App";

/* global document, Office, module, require */

initializeIcons();

let isOfficeInitialized = false;

const title = "AI Assistant";

const container = document.getElementById("container");
const root = createRoot(container);

const render = (Component) => {
    root.render(
        <AppContainer>
            <ThemeProvider>
                <Component title={title} isOfficeInitialized={isOfficeInitialized} />
            </ThemeProvider>
        </AppContainer>
    );
};

/* Render application after Office initializes */
Office.onReady(() => {
    isOfficeInitialized = true;
    render(App);
});

if ((module as any).hot) {
    (module as any).hot.accept("./components/App", () => {
        const NextApp = require("./components/App").default;
        render(NextApp);
    });
}