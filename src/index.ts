import { LoadingDialog, waitForTheme } from "dattatable";
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
    title?: string;
    sourceUrl?: string;
}

// Create the global variable for this solution
const GlobalVariable = {
    App: null,
    Configuration,
    description: Strings.ProjectDescription,
    render: (props: IProps) => {
        // Show a loading dialog
        LoadingDialog.setHeader("Loading Application");
        LoadingDialog.setBody("This may take time based on the number of apps to load...");
        LoadingDialog.show();

        // See if the page context exists
        if (props.context) {
            // Set the context
            setContext(props.context, props.envType, props.sourceUrl);

            // Update the configuration
            Configuration.setWebUrl(props.sourceUrl || ContextInfo.webServerRelativeUrl);
        }

        // Update the ProjectName from SPFx title field
        props.title ? Strings.ProjectName = props.title : null;

        // Initialize the application
        DataSource.init().then(
            // Success
            () => {
                // Update the loading dialog
                LoadingDialog.setHeader("Loading Theme");

                // Wait for the theme to be loaded
                waitForTheme().then(() => {
                    // Create the application
                    GlobalVariable.App = new App(props.el);
                    
                    // Hide the loading dialog
                    LoadingDialog.hide();
                });
            },

            // Error
            () => {
                // Display the install dialog
                InstallationModal.show();

                // Hide the loading dialog
                LoadingDialog.hide();
            }
        );
    },
    setAppCatalogUrl: (url: string) => {
        // Set the app catalog url
        DataSource.AppCatalogUrl = url;
    },
    title: Strings.ProjectName,
    updateTheme: (themeInfo) => {
        // Set the theme
        ThemeManager.setCurrentTheme(themeInfo);
    }
};

// Make is available in the DOM
window[Strings.GlobalVariable] = GlobalVariable;

// Get the element and render the app if it is found
let elApp = document.querySelector("#" + Strings.AppElementId) as HTMLElement;
if (elApp) {
    // Remove the extra border spacing on the webpart in classic mode
    let contentBox = document.querySelector("#contentBox table.ms-core-tableNoSpace");
    contentBox ? contentBox.classList.remove("ms-webpartPage-root") : null;

    // Set the app catalog url property
    DataSource.AppCatalogUrl = elApp.getAttribute("data-appCatalogUrl");

    // Render the application
    GlobalVariable.render({ el: elApp });
}