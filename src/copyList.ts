import { List, LoadingDialog, Modal } from "dattatable";
import { Components, ContextInfo, Helper, SPTypes, Types, Web } from "gd-sprest-bs";
import { IAppStoreItem } from "./ds";
import Strings from "./strings";

/**
 * Copy List Modal
 */
export class CopyListModal {
    // Method to copy the list
    private static copyList(dstListName: string, srcWebUrl: string, srcList: Types.SP.List): PromiseLike<void> {
        // Show a loading dialog
        LoadingDialog.setHeader("Copying the List");
        LoadingDialog.setBody("Initializing the request...");
        LoadingDialog.show();

        // Return a promise
        return new Promise((resolve, reject) => {
            // Ensure the user has the correct permissions to create the list
            let dstWebUrl = Strings.ListTemplateUrl
                .replace("~sitecollection", ContextInfo.siteServerRelativeUrl)
                .replace("~site", ContextInfo.webServerRelativeUrl);
            Web(dstWebUrl).query({ Expand: ["EffectiveBasePermissions"] }).execute(
                // Exists
                web => {
                    // Ensure the user doesn't have permission to manage lists
                    if (!Helper.hasPermissions(web.EffectiveBasePermissions, [SPTypes.BasePermissionTypes.ManageLists])) {
                        // Reject the request
                        reject("You do not have permission to create lists on this web.");
                        LoadingDialog.hide();
                        return;
                    }

                    // Getting the source list information
                    LoadingDialog.setBody("Loading source list...");

                    // Get the list information
                    var list = new List({
                        listName: srcList.Title,
                        webUrl: srcWebUrl,
                        itemQuery: { Filter: "Id eq 0" },
                        onInitError: () => {
                            // Reject the request
                            reject("Error loading the list information. Please check your permission to the source list.");
                            LoadingDialog.hide();
                            return;
                        },
                        onInitialized: () => {
                            // Update the loading dialog
                            LoadingDialog.setBody("Analyzing the list information...");

                            // Create the configuration
                            let cfgProps: Helper.ISPConfigProps = {
                                ListCfg: [{
                                    ListInformation: {
                                        BaseTemplate: list.ListInfo.BaseTemplate,
                                        Title: dstListName,
                                        AllowContentTypes: list.ListInfo.AllowContentTypes,
                                        Hidden: list.ListInfo.Hidden,
                                        NoCrawl: list.ListInfo.NoCrawl
                                    },
                                    CustomFields: [],
                                    ViewInformation: []
                                }]
                            };

                            // Parse the content type fields
                            let lookupFields: Types.SP.FieldLookup[] = [];
                            for (let i = 0; i < list.ListContentTypes[0].Fields.results.length; i++) {
                                let fldInfo = list.ListContentTypes[0].Fields.results[i];

                                // Skip internal fields
                                if (fldInfo.InternalName == "ContentType" || fldInfo.InternalName == "Title") { continue; }

                                // See if this is a lookup field
                                if (fldInfo.FieldTypeKind == SPTypes.FieldType.Lookup) {
                                    // Add the field
                                    lookupFields.push(fldInfo);
                                }

                                // Add the field information
                                cfgProps.ListCfg[0].CustomFields.push({
                                    name: fldInfo.InternalName,
                                    schemaXml: fldInfo.SchemaXml
                                });
                            }

                            // Parse the views
                            for (let i = 0; i < list.ListViews.length; i++) {
                                let viewInfo = list.ListViews[i];

                                // Add the view
                                cfgProps.ListCfg[0].ViewInformation.push({
                                    Default: true,
                                    ViewName: viewInfo.Title,
                                    ViewFields: viewInfo.ViewFields.Items.results,
                                    ViewQuery: viewInfo.ViewQuery
                                });
                            }

                            // Update the loading dialog
                            LoadingDialog.setBody("Creating the destination list...");

                            // Create the list
                            let cfg = Helper.SPConfig(cfgProps);
                            cfg.setWebUrl(web.ServerRelativeUrl);
                            cfg.install().then(() => {
                                // Update the loading dialog
                                LoadingDialog.setBody("Validating the lookup field(s)...");

                                // Parse the lookup fields
                                Helper.Executor(lookupFields, lookupField => {
                                    // Return a promise
                                    return new Promise(resolve => {
                                        // Get the lookup field source list
                                        Web(srcWebUrl).Lists().getById(lookupField.LookupList).execute(srcList => {
                                            // Get the lookup list in the destination site
                                            Web(web.ServerRelativeUrl).Lists(srcList.Title).execute(dstList => {
                                                // Update the field schema xml
                                                let fieldDef = lookupField.SchemaXml.replace(`List="${lookupField.LookupList}"`, `List="{${dstList.Id}}"`);
                                                Web(web.ServerRelativeUrl).Lists(list.ListInfo.Title).Fields(lookupField.InternalName).update({
                                                    SchemaXml: fieldDef
                                                }).execute(() => {
                                                    // Updated the lookup list
                                                    console.log(`Updated the lookup field '${lookupField.InternalName}' in lookup list successfully.`);

                                                    // Check the next field
                                                    resolve(null);
                                                })
                                            }, () => {
                                                // Error getting the lookup list
                                                console.error(`Error getting the lookup list '${lookupField.LookupList}' from web '${web.ServerRelativeUrl}'.`);

                                                // Check the next field
                                                resolve(null);
                                            });
                                        }, () => {
                                            // Error getting the lookup list
                                            console.error(`Error getting the lookup list '${lookupField.LookupList}' from web '${srcWebUrl}'.`);

                                            // Check the next field
                                            resolve(null);
                                        });
                                    });
                                }).then(() => {
                                    // Hide the loading dialog
                                    LoadingDialog.hide();

                                    // Resolve the request
                                    resolve();
                                });
                            });
                        }
                    });
                },

                // Doesn't exist
                () => {
                    // Hide the loading dialog
                    LoadingDialog.hide();

                    // Reject the request
                    reject("Error getting the target web. You may not have permissions or it doesn't exist.");
                }
            );
        });
    }

    // Displays the modal
    static show(appItem: IAppStoreItem, webUrl: string = ContextInfo.webServerRelativeUrl) {
        // Clear the modal
        Modal.clear();

        // Set the header
        Modal.setHeader("Copy List");

        // Render a form
        let form = Components.Form({
            el: Modal.BodyElement,
            controls: [
                {
                    name: "WebUrl",
                    title: "Source Web Url",
                    type: Components.FormControlTypes.TextField,
                    description: "The source web containing the list.",
                    required: true,
                    errorMessage: "A relative web url is required. (Ex. /sites/dev)",
                    value: webUrl
                },
                {
                    name: "SourceList",
                    title: "Select a List",
                    type: Components.FormControlTypes.Dropdown,
                    description: "Select the list to copy.",
                    required: true,
                    errorMessage: "A list is required."
                }
            ]
        });

        // Set the footer
        Components.TooltipGroup({
            el: Modal.FooterElement,
            tooltips: [
                {
                    content: "Loads the lists from the selected web.",
                    btnProps: {
                        text: "Load Lists",
                        type: Components.ButtonTypes.OutlinePrimary,
                        onClick: () => {
                            // Get the web url control
                            let ctrlWeb = form.getControl("WebUrl");

                            // Validate the control
                            if (ctrlWeb.isValid) {
                                // Get the dropdown control
                                let ctrlLists = form.getControl("SourceList");

                                // Show a loading dialog
                                LoadingDialog.setHeader("Loading Lists");
                                LoadingDialog.setBody("This dialog will close after the lists are loaded.");
                                LoadingDialog.show();

                                // Load the lists
                                Web().Lists().execute(
                                    // Success
                                    lists => {
                                        let items: Components.IDropdownItem[] = [];

                                        // Parse the lists
                                        for (let i = 0; i < lists.results.length; i++) {
                                            let list = lists.results[i];

                                            // Add the item
                                            items.push({
                                                data: list,
                                                text: list.Title,
                                                value: list.Title
                                            });
                                        }

                                        // Set the dropdowns
                                        ctrlLists.dropdown.setItems(items);

                                        // Hide the loading dialog
                                        LoadingDialog.hide();
                                    },

                                    // Error
                                    () => {
                                        // Clear the dropdown items
                                        ctrlLists.dropdown.setItems([]);

                                        // Set the validation
                                        ctrlLists.updateValidation(ctrlLists.el, {
                                            isValid: false,
                                            invalidMessage: "Error loading the lists from the web."
                                        });

                                        // Hide the loading dialog
                                        LoadingDialog.hide();
                                    }
                                );
                            }
                        }
                    }
                },
                {
                    content: "Copies the selected list to the list templates sub-web.",
                    btnProps: {
                        text: "Copy List",
                        type: Components.ButtonTypes.OutlinePrimary,
                        onClick: () => {
                            // Validate the form
                            if (form.isValid()) {
                                let ctrlLists = form.getControl("SourceList");
                                let formValues = form.getValues();

                                // Set the list name
                                let listData = formValues["SourceList"].data as Types.SP.List;
                                let listName = `${appItem.Title} ${listData.Title}`;

                                // Copy the list
                                this.copyList(listName, formValues["WebUrl"], listData).then(
                                    () => {
                                        // See if the associated list doesn't exist
                                        let lists = (appItem.AssociatedLists || "").split('\n');
                                        if (lists.indexOf(listName) < 0) {
                                            // Append the list
                                            lists.push(listName);

                                            // Update the item
                                            appItem.update({
                                                AssociatedLists: lists.join('\n')
                                            }).execute();
                                        }

                                        // Update the validation
                                        ctrlLists.updateValidation(ctrlLists.el, {
                                            isValid: true,
                                            validMessage: "List copied successfully."
                                        });
                                    },
                                    err => {
                                        // Show the error message
                                        ctrlLists.updateValidation(ctrlLists.el, {
                                            isValid: false,
                                            invalidMessage: err
                                        });
                                    }
                                );
                            }
                        }
                    }
                },
                {
                    content: "Closes the dialog.",
                    btnProps: {
                        text: "Close",
                        type: Components.ButtonTypes.OutlineDanger,
                        onClick: () => {
                            // Close the dialog
                            Modal.hide();
                        }
                    }
                }
            ]
        });

        // Show the modal
        Modal.show();
    }
}
