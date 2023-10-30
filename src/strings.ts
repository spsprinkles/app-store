import { ContextInfo } from "gd-sprest-bs";

// Sets the context information
// This is for SPFx or Teams solutions
export const setContext = (context, sourceUrl?: string) => {
    // Set the context
    ContextInfo.setPageContext(context.pageContext);

    // Update the source url
    Strings.SourceUrl = sourceUrl || ContextInfo.webServerRelativeUrl;
}

// Gets the list template url
export const getListTemplateUrl = () => {
    return Strings.ListTemplateUrl
        .replace("~sitecollection", ContextInfo.siteServerRelativeUrl)
        .replace("~site", ContextInfo.webServerRelativeUrl);
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
    ListTemplateUrl: "~site/listtemplates",
    ProjectName: "AppStore",
    ProjectDescription: "List that displays the app information in a card view.",
    SecurityGroups: {
        Developers: {
            Name: "App Store Developers",
            Description: "Contributors to the app store solution."
        },
        Managers: {
            Name: "App Store Managers",
            Description: "Manages the app store solution."
        }
    },
    SourceUrl: ContextInfo.webServerRelativeUrl,
    Version: "0.1.0"
};
export default Strings;