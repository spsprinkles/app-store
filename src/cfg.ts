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
            CustomFields: [
                
                {
                    name: "ProjectName",
                    title: "Project name",
                    type: Helper.SPCfgFieldType.Text,
                    description: "Name of the project",
                    defaultValue: "",
                    required: true,

                },

                {
                    name: "Description",
                    title: "Description",
                    type: Helper.SPCfgFieldType.Text,
                    description: "Description of the project",
                    defaultValue: "",
                    required: true,
                },

                {
                    name: "AppScreenShot",
                    title: "App Screenshot",
                    type: Helper.SPCfgFieldType.Choice,
                    defaultValue: "",
                    required: true,
                    showInNewForm: false,
                    choices: [
                        "Screenshot 1",
                        "Screenshot 2",
                        "Screenshot 3",
                        "Screenshot 4",
                        "Screenshot 5",
                    ],
                } as Helper.IFieldInfoChoice,

                {
                    name: "AppVideoURL",
                    title: "App Video URL",
                    type: Helper.SPCfgFieldType.Text,
                    defaultValue: "",
                    required: true,
                },
            ],
            ViewInformation: [
                {
                    ViewName: "All Items",
                    ViewFields: [
                        "ProjectName", "Description", "AppScreenShot", "AppVideoURL",
                    ]
                }
            ]
        }
    ]
});