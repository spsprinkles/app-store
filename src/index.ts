import { ContextInfo, ThemeManager } from "gd-sprest-bs";
import { App } from "./app";
import { Configuration } from "./cfg";
import { DataSource } from "./ds";
import { InstallationModal } from "./install";
import Strings, { setContext } from "./strings";

// Styling
import "./styles.scss";

// Properties
interface IProps {
    el: HTMLElement;
    context?: any;
    displayMode?: number;
    envType?: number;
    sourceUrl?: string;
}

// Create the global variable for this solution
const GlobalVariable = {
    App: null,
    Configuration,
    description: Strings.ProjectDescription,
    render: (props: IProps) => {
        // See if the page context exists
        if (props.context) {
            // Set the context
            setContext(props.context, props.envType, props.sourceUrl);

            // Update the configuration
            Configuration.setWebUrl(props.sourceUrl || ContextInfo.webServerRelativeUrl);
        }

        // Initialize the application
        DataSource.init().then(
            // Success
            () => {
                // Load the theme
                ThemeManager.load(true).then(() => {
                    // Create the application
                    GlobalVariable.App = new App(props.el);
                });
            },

            // Error
            () => {
                // Display the install dialog
                InstallationModal.show();
            }
        );
    },
    setAppCatalogUrl: (url: string) => {
        // Set the app catalog url
        DataSource.AppCatalogUrl = url;
    },
    updateTheme: (themeInfo) => {
        // Set the theme
        ThemeManager.setCurrentTheme(themeInfo);
    },
    version: Strings.Version
};

// Make is available in the DOM
window[Strings.GlobalVariable] = GlobalVariable;

// Get the element and render the app if it is found
let elApp = document.querySelector("#" + Strings.AppElementId) as HTMLElement;
if (elApp) {
    // Set the app catalog url property
    DataSource.AppCatalogUrl = elApp.getAttribute("data-appCatalogUrl");

    // Render the application
    GlobalVariable.render({ el: elApp });
}