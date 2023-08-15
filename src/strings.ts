import { ContextInfo } from "gd-sprest-bs";

// Sets the context information
// This is for SPFx or Teams solutions
export const setContext = (context, sourceUrl?: string) => {
    // Set the context
    ContextInfo.setPageContext(context.pageContext);

    // Update the source url
    Strings.SourceUrl = sourceUrl || ContextInfo.webServerRelativeUrl;
}

/**
 * Global Constants
 */
const Strings = {
    AppElementId: "app-store",
    DateFormat: "YYYY-MMM-DD",
    GlobalVariable: "AppStore",
    Lists: {
        Main: "App Store",
        Ratings: "App Ratings"
    },
    ProjectName: "AppStore",
    ProjectDescription: "List that displays the app information in a card view.",
    SecurityGroups: {
        Managers: {
            Name: "App Store Managers",
            Description: "Manages the app store solution."
        }
    },
    SourceUrl: ContextInfo.webServerRelativeUrl,
    Version: "0.0.6"
};
export default Strings;