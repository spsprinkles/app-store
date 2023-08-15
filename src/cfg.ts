import { Helper, SPTypes } from "gd-sprest-bs";
import Strings from "./strings";

/**
 * SharePoint Assets
 */
export const Configuration = Helper.SPConfig({
    ListCfg: [
        {
            ListInformation: {
                Title: Strings.Lists.Main,
                BaseTemplate: SPTypes.ListTemplateType.GenericList,
                ContentTypesEnabled: true
            },
            TitleFieldDisplayName: "App Name",
            ContentTypes: [
                {
                    Name: "Item",
                    FieldRefs: [
                        "Title",
                        "AppType",
                        "Status",
                        "Description",
                        "Developers",
                        "Organization",
                        "MoreInfo",
                        "SupportURL",
                        "Icon",
                        "ScreenShot1",
                        "ScreenShot2",
                        "ScreenShot3",
                        "ScreenShot4",
                        "ScreenShot5",
                        "VideoURL"
                    ]
                }
            ],
            CustomFields: [
                {
                    name: "AppType",
                    title: "App Type",
                    type: Helper.SPCfgFieldType.Choice,
                    required: true,
                    choices: [
                        "Power Apps", "Power Automate", "Power BI", "PowerShell", "SharePoint", "Teams"
                    ]
                } as Helper.IFieldInfoChoice,
                {
                    name: "Description",
                    title: "Description",
                    type: Helper.SPCfgFieldType.Note,
                    description: "A description of the application and its function and usage",
                    notetype: SPTypes.FieldNoteType.TextOnly,
                    required: true,
                } as Helper.IFieldInfoNote,
                {
                    name: "Developers",
                    title: "Developers",
                    type: Helper.SPCfgFieldType.User,
                    description: "The developers of the application",
                    multi: true,
                    required: true,
                    selectionMode: SPTypes.FieldUserSelectionType.PeopleOnly,
                    showField: "ImnName"
                } as Helper.IFieldInfoUser,
                {
                    name: "Icon",
                    title: "Icon",
                    type: Helper.SPCfgFieldType.Note,
                    description: "A custom icon for the application",
                    notetype: SPTypes.FieldNoteType.TextOnly
                } as Helper.IFieldInfoNote,
                {
                    name: "MoreInfo",
                    title: "More Info",
                    type: Helper.SPCfgFieldType.Url,
                    description: "A link for more information, or to launch the application"
                } as Helper.IFieldInfoUrl,
                {
                    name: "Organization",
                    title: "Organization",
                    type: Helper.SPCfgFieldType.Text,
                    description: "The organization that supports the application",
                    required: true
                },
                {
                    name: "Rating",
                    title: "Rating",
                    type: Helper.SPCfgFieldType.Number,
                    defaultValue: "0",
                    min: 0,
                    max: 5,
                    numberType: SPTypes.FieldNumberType.Integer
                } as Helper.IFieldInfoNumber,
                {
                    name: "RatingCount",
                    title: "Rating Count",
                    type: Helper.SPCfgFieldType.Number,
                    defaultValue: "0",
                    numberType: SPTypes.FieldNumberType.Integer
                } as Helper.IFieldInfoNumber,
                {
                    name: "ScreenShot1",
                    title: "Screen Shot 1",
                    type: Helper.SPCfgFieldType.Note,
                    description: "A screenshot of the application",
                    notetype: SPTypes.FieldNoteType.TextOnly,
                    required: true
                } as Helper.IFieldInfoNote,
                {
                    name: "ScreenShot2",
                    title: "Screen Shot 2",
                    type: Helper.SPCfgFieldType.Note,
                    description: "A screenshot of the application",
                    notetype: SPTypes.FieldNoteType.TextOnly
                } as Helper.IFieldInfoNote,
                {
                    name: "ScreenShot3",
                    title: "Screen Shot 3",
                    type: Helper.SPCfgFieldType.Note,
                    description: "A screenshot of the application",
                    notetype: SPTypes.FieldNoteType.TextOnly
                } as Helper.IFieldInfoNote,
                {
                    name: "ScreenShot4",
                    title: "Screen Shot 4",
                    type: Helper.SPCfgFieldType.Note,
                    description: "A screenshot of the application",
                    notetype: SPTypes.FieldNoteType.TextOnly
                } as Helper.IFieldInfoNote,
                {
                    name: "ScreenShot5",
                    title: "Screen Shot 5",
                    type: Helper.SPCfgFieldType.Note,
                    description: "A screenshot of the application",
                    notetype: SPTypes.FieldNoteType.TextOnly
                } as Helper.IFieldInfoNote,
                {
                    name: "Status",
                    title: "Status",
                    type: Helper.SPCfgFieldType.Choice,
                    defaultValue: "New",
                    required: true,
                    showInNewForm: false,
                    choices: [
                        "New", "In Testing", "Approved", "Depricated"
                    ]
                } as Helper.IFieldInfoChoice,
                {
                    name: "SupportURL",
                    title: "Support URL",
                    type: Helper.SPCfgFieldType.Url,
                    description: "A link to a support page"
                } as Helper.IFieldInfoUrl,
                {
                    name: "VideoURL",
                    title: "App Video URL",
                    type: Helper.SPCfgFieldType.Url,
                    defaultValue: "",
                    description: "A link to a video demo of the application"
                } as Helper.IFieldInfoUrl,
            ],
            ViewInformation: [
                {
                    ViewName: "All Items",
                    ViewFields: [
                        "LinkTitle", "AppType", "Status", "Description", "Developers", "Organization", "MoreInfo", "SupportURL"
                    ]
                }
            ]
        },
        {
            ListInformation: {
                Title: Strings.Lists.Ratings,
                BaseTemplate: SPTypes.ListTemplateType.GenericList
            },
            ContentTypes: [
                {
                    Name: "Item",
                    FieldRefs: [
                        "Title",
                        "AppLU",
                        "Rating",
                        "Comment"
                    ]
                }
            ],
            CustomFields: [
                {
                    name: "AppLU",
                    title: "App",
                    type: Helper.SPCfgFieldType.Lookup,
                    indexed: true,
                    listName: Strings.Lists.Main,
                    required: true,
                    showField: "Title"
                } as Helper.IFieldInfoLookup,
                {
                    name: "Comment",
                    title: "Comment",
                    type: Helper.SPCfgFieldType.Text
                },
                {
                    name: "Rating",
                    title: "Rating",
                    type: Helper.SPCfgFieldType.Number,
                    min: 0,
                    max: 5,
                    numberType: SPTypes.FieldNumberType.Integer
                } as Helper.IFieldInfoNumber
            ]
        }
    ]
});