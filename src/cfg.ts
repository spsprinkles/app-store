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
                },
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
                    type: Helper.SPCfgFieldType.Url
                },
                {
                    name: "ScreenShot1",
                    title: "Screen Shot 1",
                    type: Helper.SPCfgFieldType.Url,
                    required: true,
                },
                {
                    name: "ScreenShot2",
                    title: "Screen Shot 2",
                    type: Helper.SPCfgFieldType.Url
                },
                {
                    name: "ScreenShot3",
                    title: "Screen Shot 3",
                    type: Helper.SPCfgFieldType.Url
                },
                {
                    name: "ScreenShot4",
                    title: "Screen Shot 4",
                    type: Helper.SPCfgFieldType.Url
                },
                {
                    name: "ScreenShot5",
                    title: "Screen Shot 5",
                    type: Helper.SPCfgFieldType.Url
                },
                {
                    name: "TypeOfProject",
                    title: "Type of Project",
                    type: Helper.SPCfgFieldType.Choice,
                    required: true,
                    choices: [
                        "Power App", "Power Automate", "Power Shell", "SharePoint", "Teams"
                    ]
                } as Helper.IFieldInfoChoice,
                {
                    name: "VideoURL",
                    title: "App Video URL",
                    type: Helper.SPCfgFieldType.Url,
                    defaultValue: "",
                    required: false,
                },
            ],
            ViewInformation: [
                {
                    ViewName: "All Items",
                    ViewFields: [
                        "LinkTitle", "Description"
                    ]
                }
            ]
        }
    ]
});