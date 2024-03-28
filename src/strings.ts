import { ContextInfo, SPTypes } from "gd-sprest-bs";

// Sets the context information
// This is for SPFx or Teams solutions
export const setContext = (context, envType?: number, sourceUrl?: string) => {
    // Set the context
    ContextInfo.setPageContext(context.pageContext);

    // Update the properties
    Strings.IsClassic = envType == SPTypes.EnvironmentType.ClassicSharePoint;
    Strings.IsFlow3 = ContextInfo.blockDownloadsExperienceEnabled ? true : false;
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
    IsClassic: true,
    IsFlow3: ContextInfo.blockDownloadsExperienceEnabled ? true : false,
    Lists: {
        Main: "App Store",
        Ratings: "App Ratings",
        Requests: "App Store Requests"
    },
    ListTemplateWebInfo: {
        Description: "Stores the list templates for the app store solution.",
        Title: "List Templates",
        Url: "listtemplates",
        WebTemplate: SPTypes.WebTemplateType.Site
    },
    ListTemplateUrl: "~site/listtemplates",
    ProjectName: "Solution Center",
    ProjectDescription: "An app that displays solutions in a card view.",
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
    Version: "0.1.5"
};
export default Strings;