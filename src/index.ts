import { ContextInfo } from "gd-sprest-bs";
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
                // Create the application
                GlobalVariable.App = new App(props.el);

                if (Strings.IsClassic) {
                    let counter = 0;
                    let loopId = setInterval(() => {
                        // See if the theme exists
                        if (ContextInfo.theme.accent) {
                            clearInterval(loopId);
                            console.log("It took " + counter + " tries for the theme to exist.");
                            GlobalVariable.updateTheme(ContextInfo.theme);
                        } else if (++counter > 10) {
                            clearInterval(loopId);
                        }
                    }, 100);
                }
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
        // Store the theme info
        DataSource.ThemeInfo = themeInfo;

        // See if the app exists
        if (GlobalVariable.App) {
            // Apply theming
            GlobalVariable.App.updateTheme();
        }
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