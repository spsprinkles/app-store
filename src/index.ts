import { ContextInfo } from "gd-sprest-bs";
import { App } from "./app";
import { Configuration } from "./cfg";
import { DataSource } from "./ds";
import { InstallationModal } from "./install";
import Strings, { setContext } from "./strings";

// Styling
import "./styles.scss";

// Create the global variable for this solution
const GlobalVariable = {
    App: null,
    Configuration,
    description: Strings.ProjectDescription,
    render: (el, context?, sourceUrl?: string) => {
        // See if the page context exists
        if (context) {
            // Set the context
            setContext(context, sourceUrl);

            // Update the configuration
            Configuration.setWebUrl(sourceUrl || ContextInfo.webServerRelativeUrl);
        }

        // Initialize the application
        DataSource.init().then(
            // Success
            () => {
                // Create the application
                GlobalVariable.App = new App(el);
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
        // TODO
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
    GlobalVariable.render(elApp);
}