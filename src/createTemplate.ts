import { List, LoadingDialog, Modal } from "dattatable";
import { Components, Helper, SPTypes, Types, Web } from "gd-sprest-bs";
import { IAppStoreItem } from "./ds";
import { getListTemplateUrl } from "./strings";

/**
 * Create Templates
 */
export class CreateTemplate {
    private static _form: Components.IForm = null;

    // Method to copy the list
    private static copyList(appItem: IAppStoreItem, srcWebUrl: string, srcList: Types.SP.List): PromiseLike<string> {
        // Show a loading dialog
        LoadingDialog.setHeader("Copying the List");
        LoadingDialog.setBody("Initializing the request...");
        LoadingDialog.show();

        // Set the destination list name
        let dstListName = `${appItem.Title} ${srcList.Title}`;

        // Return a promise
        return new Promise((resolve, reject) => {
            // Ensure the user has the correct permissions to create the list
            let dstWebUrl = getListTemplateUrl();
            Web(dstWebUrl).query({ Expand: ["EffectiveBasePermissions"] }).execute(
                // Exists
                web => {
                    // Ensure the user doesn't have permission to manage lists
                    if (!Helper.hasPermissions(web.EffectiveBasePermissions, [SPTypes.BasePermissionTypes.ManageLists])) {
                        // Reject the request
                        reject("You do not have permission to copy templates on this web.");
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
                        },
                        onInitialized: () => {
                            let lookupFields: Types.SP.FieldLookup[] = [];

                            // Update the loading dialog
                            LoadingDialog.setBody("Analyzing the list information...");

                            // Create the configuration
                            let cfgProps: Helper.ISPConfigProps = {
                                ContentTypes: [],
                                ListCfg: [{
                                    ListInformation: {
                                        AllowContentTypes: list.ListInfo.AllowContentTypes,
                                        BaseTemplate: list.ListInfo.BaseTemplate,
                                        ContentTypesEnabled: list.ListInfo.ContentTypesEnabled,
                                        Title: dstListName,
                                        Hidden: list.ListInfo.Hidden,
                                        NoCrawl: list.ListInfo.NoCrawl
                                    },
                                    ContentTypes: [],
                                    CustomFields: [],
                                    ViewInformation: []
                                }]
                            };

                            // Parse the content types
                            for (let i = 0; i < list.ListContentTypes.length; i++) {
                                let ct = list.ListContentTypes[i];

                                // Skip sealed content types
                                if (ct.Sealed) { continue; }

                                // Skip the internal content types
                                if (ct.Name != "Document" && ct.Name != "Event" && ct.Name != "Item" && ct.Name != "Task") {
                                    // Add the content type
                                    cfgProps.ContentTypes.push({
                                        Name: ct.Name,
                                        ParentName: "Item"
                                    });
                                }

                                // Parse the content type fields
                                let fieldRefs = [];
                                for (let j = 0; j < ct.Fields.results.length; j++) {
                                    let fldInfo = ct.Fields.results[j];

                                    // Skip internal fields
                                    if (fldInfo.InternalName == "ContentType" || fldInfo.InternalName == "Title") { continue; }

                                    // Append the field ref
                                    fieldRefs.push(fldInfo.InternalName);

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

                                // Add the list content type
                                cfgProps.ListCfg[0].ContentTypes.push({
                                    Name: ct.Name,
                                    Description: ct.Description,
                                    ParentName: ct.Name,
                                    FieldRefs: fieldRefs
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

                            // Save the configuration as a string
                            let strConfig = JSON.stringify(cfgProps);

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
                                        Web(srcWebUrl).Lists().getById(lookupField.LookupList).execute(srcLookupList => {
                                            // The lookup list will need to be copied first and should have the same
                                            // format of the template "[App Title] [List Title]"
                                            // Set the lookup list name to match
                                            let dstLookupList = `${appItem.Title} ${srcLookupList.Title}`;

                                            // Get the lookup list in the destination site
                                            Web(web.ServerRelativeUrl).Lists(dstLookupList).execute(dstList => {
                                                // Update the field schema xml
                                                let fieldDef = lookupField.SchemaXml.replace(`List="${lookupField.LookupList}"`, `List="{${dstList.Id}}"`);
                                                Web(web.ServerRelativeUrl).Lists(dstListName).Fields(lookupField.InternalName).update({
                                                    SchemaXml: fieldDef
                                                }).execute(() => {
                                                    // Updated the lookup list
                                                    console.log(`Updated the lookup field '${lookupField.InternalName}' in lookup list successfully.`);

                                                    // Check the next field
                                                    resolve(null);
                                                })
                                            }, () => {
                                                // Error getting the lookup list
                                                console.error(`Error getting the lookup list '${dstLookupList}' from web '${web.ServerRelativeUrl}'.`);

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
                                    // Update the list configuration
                                    this.updateListConfiguration(appItem, dstListName, JSON.parse(strConfig)).then(() => {
                                        // Hide the loading dialog
                                        LoadingDialog.hide();

                                        // Resolve the request
                                        resolve(dstListName);
                                    });
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

    // Renders the footer
    static renderFooter(el: HTMLElement, appItem: IAppStoreItem, clearFl: boolean = true) {
        // Clear the footer
        if (clearFl) { while (el.firstChild) { el.removeChild(el.firstChild); } }

        // Set the footer
        Components.TooltipGroup({
            el,
            className: "float-end",
            tooltips: [
                {
                    content: "Load the lists from the source web",
                    btnProps: {
                        text: "Load Lists",
                        type: Components.ButtonTypes.OutlinePrimary,
                        onClick: () => {
                            // Get the web url control
                            let ctrlWeb = this._form.getControl("WebUrl");

                            // Validate the control
                            if (ctrlWeb.isValid) {
                                // Get the dropdown control
                                let ctrlLists = this._form.getControl("SourceList");

                                // Show a loading dialog
                                LoadingDialog.setHeader("Loading Lists");
                                LoadingDialog.setBody("This dialog will close after the lists are loaded");
                                LoadingDialog.show();

                                // Load the lists
                                Web(ctrlWeb.getValue()).Lists().execute(
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

                                        // Set the validation
                                        ctrlLists.updateValidation(ctrlLists.el, {
                                            isValid: true
                                        });

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
                                            invalidMessage: "Error loading the lists from the web"
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
                    content: "Copy the selected list as a list template",
                    btnProps: {
                        text: "Copy List",
                        type: Components.ButtonTypes.OutlinePrimary,
                        onClick: () => {
                            // Validate the form
                            if (this._form.isValid()) {
                                let ctrlLists = this._form.getControl("SourceList");
                                let formValues = this._form.getValues();

                                // Set the list name
                                let listData = formValues["SourceList"].data as Types.SP.List;

                                // Copy the list
                                this.copyList(appItem, formValues["WebUrl"], listData).then(
                                    (listName: string) => {
                                        // Update the validation
                                        ctrlLists.updateValidation(ctrlLists.el, {
                                            isValid: true,
                                            validMessage: "List copied successfully"
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
                }
            ]
        });
    }

    // Renders the main form
    static renderForm(el: HTMLElement, webUrl?: string) {
        // Set the body
        el.innerHTML = `<label class="my-2">Use this form to add or update a list template for this app.`;

        // Render a form
        this._form = Components.Form({
            el,
            controls: [
                {
                    name: "WebUrl",
                    title: "Source Web Url",
                    type: Components.FormControlTypes.TextField,
                    description: "The source web containing the list",
                    required: true,
                    errorMessage: "A relative web url is required (Ex. /sites/dev)",
                    value: webUrl
                },
                {
                    name: "SourceList",
                    title: "Select a List",
                    type: Components.FormControlTypes.Dropdown,
                    description: "Select a list to copy",
                    required: true,
                    errorMessage: "A list is required"
                }
            ]
        });
    }

    // Updates the list configuration for the item
    private static updateListConfiguration(appItem: IAppStoreItem, listName: string, cfg: Helper.ISPConfigProps): PromiseLike<void> {
        // Return a promise
        return new Promise(resolve => {
            // See if the associated list doesn't exist
            let associatedLists = (appItem.AssociatedLists || "").split('\n');
            if (associatedLists.indexOf(listName) < 0) {
                // Append the list
                associatedLists.push(listName);
            }

            // Get the list configurations and append/replace it
            let listCfg = appItem.ListConfigurations;
            try {
                // Get the configuration
                let appConfig: Helper.ISPConfigProps = JSON.parse(listCfg);

                // See if content types exist
                if (cfg.ContentTypes?.length > 0) {
                    // Ensure content types are defined
                    appConfig.ContentTypes = appConfig.ContentTypes || [];

                    // Parse the content types
                    let newCTS = cfg.ContentTypes || [];
                    for (let i = 0; i < newCTS.length; i++) {
                        let foundFl = false;
                        let newCT = newCTS[i];

                        // Parse the app configuration
                        for (let j = 0; j < appConfig.ContentTypes.length; j++) {
                            let ct = appConfig.ContentTypes[j];

                            // See if they match
                            if (newCT.Name == ct.Name) {
                                // Replace it
                                appConfig.ContentTypes[j] = newCT;
                                foundFl = true;
                                break;
                            }
                        }

                        // See if it wasn't found
                        if (!foundFl) {
                            // Append it
                            appConfig.ContentTypes.push(newCT)
                        }
                    }
                }

                // Parse the lists
                let foundFl = false;
                for (let i = 0; i < appConfig.ListCfg.length; i++) {
                    // See if this is the target list
                    if (appConfig.ListCfg[i].ListInformation.Title == listName) {
                        // Replace it
                        appConfig.ListCfg[i] = cfg.ListCfg[0];
                        foundFl = true;
                        break;
                    }
                }

                // See if it wasn't found
                if (!foundFl) {
                    // Append the configuration
                    appConfig.ListCfg.push(cfg.ListCfg[0]);
                }

                // Set the configuration
                listCfg = JSON.stringify(appConfig);
            } catch {
                // Set the configuration
                listCfg = JSON.stringify(cfg);
            }

            // Update the item
            appItem.update({
                AssociatedLists: associatedLists.join('\n'),
                ListConfigurations: listCfg
            }).execute(resolve);
        });
    }
}