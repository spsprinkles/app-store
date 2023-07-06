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
                BaseTemplate: SPTypes.ListTemplateType.GenericList
            },
            ContentTypes: [
                {
                    Name: "Item",
                    FieldRefs: [
                        "Title",
                        "TypeOfProject",
                        "Description",
                        "AdditionalInformation",
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
                    name: "AdditionalInformation",
                    title: "Additional Information",
                    type: Helper.SPCfgFieldType.Url
                } as Helper.IFieldInfoUrl,
                {
                    name: "Description",
                    title: "Description",
                    type: Helper.SPCfgFieldType.Note,
                    description: "Description of the app/solution",
                    notetype: SPTypes.FieldNoteType.TextOnly,
                    required: true,
                } as Helper.IFieldInfoNote,
                {
                    name: "Icon",
                    title: "Icon",
                    type: Helper.SPCfgFieldType.Note,
                    description: "The icon for the app/solution.",
                    notetype: SPTypes.FieldNoteType.TextOnly
                } as Helper.IFieldInfoNote,
                {
                    name: "ScreenShot1",
                    title: "Screen Shot 1",
                    type: Helper.SPCfgFieldType.Note,
                    description: "A screenshot of the app/solution.",
                    notetype: SPTypes.FieldNoteType.TextOnly,
                    required: true
                } as Helper.IFieldInfoNote,
                {
                    name: "ScreenShot2",
                    title: "Screen Shot 2",
                    type: Helper.SPCfgFieldType.Note,
                    description: "A screenshot of the app/solution.",
                    notetype: SPTypes.FieldNoteType.TextOnly
                } as Helper.IFieldInfoNote,
                {
                    name: "ScreenShot3",
                    title: "Screen Shot 3",
                    type: Helper.SPCfgFieldType.Note,
                    description: "A screenshot of the app/solution.",
                    notetype: SPTypes.FieldNoteType.TextOnly
                } as Helper.IFieldInfoNote,
                {
                    name: "ScreenShot4",
                    title: "Screen Shot 4",
                    type: Helper.SPCfgFieldType.Note,
                    description: "A screenshot of the app/solution.",
                    notetype: SPTypes.FieldNoteType.TextOnly
                } as Helper.IFieldInfoNote,
                {
                    name: "ScreenShot5",
                    title: "Screen Shot 5",
                    type: Helper.SPCfgFieldType.Note,
                    description: "A screenshot of the app/solution.",
                    notetype: SPTypes.FieldNoteType.TextOnly
                } as Helper.IFieldInfoNote,
                {
                    name: "TypeOfProject",
                    title: "Type of Project",
                    type: Helper.SPCfgFieldType.Choice,
                    required: true,
                    choices: [
                        "Power Apps", "Power Automate", "Power BI", "PowerShell", "SharePoint", "Teams"
                    ]
                } as Helper.IFieldInfoChoice,
                {
                    name: "VideoURL",
                    title: "App Video URL",
                    type: Helper.SPCfgFieldType.Url,
                    defaultValue: "",
                    required: false,
                } as Helper.IFieldInfoUrl,
            ],
            ViewInformation: [
                {
                    ViewName: "All Items",
                    ViewFields: [
                        "LinkTitle", "TypeOfProject", "Description"
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
                        "Rating"
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
                    showField: "Title"
                } as Helper.IFieldInfoLookup,
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