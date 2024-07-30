import { List, LoadingDialog } from "dattatable";
import { Components, Helper, SPTypes, Types, Web } from "gd-sprest-bs";
import { IAppStoreItem } from "./ds";
import { getListTemplateUrl } from "./strings";
import { CreateAppLists } from "./createAppLists";

/**
 * Read App Lists
 * Reads an existing solution's list and generates a configuration for it.
 */
export class ReadAppLists {
    private static _form: Components.IForm = null;

    // Method to create the list configuration
    private static createListConfiguration(appItem: IAppStoreItem, srcWebUrl: string, srcList: Types.SP.List) {
        // Show a loading dialog
        LoadingDialog.setHeader("Copying the List");
        LoadingDialog.setBody("Initializing the request...");
        LoadingDialog.show();

        // Ensure the user has the correct permissions to create the list
        let dstWebUrl = getListTemplateUrl();
        Web(dstWebUrl).query({ Expand: ["EffectiveBasePermissions"] }).execute(
            // Exists
            web => {
                // Ensure the user doesn't have permission to manage lists
                if (!Helper.hasPermissions(web.EffectiveBasePermissions, [SPTypes.BasePermissionTypes.ManageLists])) {
                    // Reject the request
                    console.error("You do not have permission to copy templates on this web.");
                    LoadingDialog.hide();
                    return;
                }

                // Generate the list configuration
                this.generateListConfiguration(srcWebUrl, srcList).then(listCfg => {
                    // Test the configuration
                    CreateAppLists.installConfiguration(listCfg.cfg, web.ServerRelativeUrl).then(lists => {
                        // Convert the configuration to a string
                        let strConfig = JSON.stringify(listCfg.cfg);

                        // Update the list configuration
                        this.updateListConfiguration(appItem, srcList.Title, JSON.parse(strConfig)).then(() => {
                            // Hide the loading dialog
                            LoadingDialog.hide();

                            // Show the results
                            CreateAppLists.showResults(listCfg.cfg, lists, true);
                        });
                    });
                });
            },

            // Doesn't exist
            () => {
                // Hide the loading dialog
                LoadingDialog.hide();

                // Reject the request
                console.error("Error getting the target web. You may not have permissions or it doesn't exist.");
            }
        );
    }

    // Generates the list configuration
    private static generateListConfiguration(srcWebUrl: string, srcList: Types.SP.List): PromiseLike<{ cfg: Helper.ISPConfigProps, lookupFields: Types.SP.FieldLookup[] }> {
        // Getting the source list information
        LoadingDialog.setBody("Loading source list...");

        // Return a promise
        return new Promise((resolve, reject) => {
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
                    let calcFields: Types.SP.Field[] = [];
                    let fields: { [key: string]: boolean } = {};
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
                                Title: srcList.Title,
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
                        for (let j = 0; j < ct.FieldLinks.results.length; j++) {
                            let fldInfo = list.getField(ct.FieldLinks.results[j].Name);

                            // See if this is a lookup field
                            if (fldInfo.FieldTypeKind == SPTypes.FieldType.Lookup) {
                                // Ensure this isn't an associated lookup field
                                if ((fldInfo as Types.SP.FieldLookup).IsDependentLookup != true) {
                                    // Append the field ref
                                    fieldRefs.push(fldInfo.InternalName);
                                }
                            } else {
                                // Append the field ref
                                fieldRefs.push(fldInfo.InternalName);
                            }

                            // Skip internal fields
                            if (fldInfo.InternalName == "ContentType" || fldInfo.InternalName == "Title") { continue; }

                            // See if this is a calculated field
                            if (fldInfo.FieldTypeKind == SPTypes.FieldType.Calculated) {
                                // Add the field and continue the loop
                                calcFields.push(fldInfo);
                                continue;
                            }

                            // Ensure the field hasn't been added
                            if (fields[fldInfo.InternalName] == null) {
                                // Add the field information
                                fields[fldInfo.InternalName] = true;
                                cfgProps.ListCfg[0].CustomFields.push({
                                    name: fldInfo.InternalName,
                                    schemaXml: fldInfo.SchemaXml
                                });

                                // See if this is a lookup field
                                if (fldInfo.FieldTypeKind == SPTypes.FieldType.Lookup) {
                                    // Add the field
                                    lookupFields.push(fldInfo);
                                }
                            }
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

                        // Parse the fields
                        for (let j = 0; j < viewInfo.ViewFields.Items.results.length; j++) {
                            let field = list.getField(viewInfo.ViewFields.Items.results[j]);

                            // Ensure the field exists
                            if (fields[field.InternalName] == null) {
                                // See if this is a calculated field
                                if (field.FieldTypeKind == SPTypes.FieldType.Calculated) {
                                    // Add the field and continue the loop
                                    calcFields.push(field);
                                } else {
                                    // Append the field
                                    fields[field.InternalName] = true;
                                    cfgProps.ListCfg[0].CustomFields.push({
                                        name: field.InternalName,
                                        schemaXml: field.SchemaXml
                                    });

                                    // See if this is a lookup field
                                    if (field.FieldTypeKind == SPTypes.FieldType.Lookup) {
                                        // Add the field
                                        lookupFields.push(field);
                                    }
                                }
                            }
                        }

                        // Add the view
                        cfgProps.ListCfg[0].ViewInformation.push({
                            Default: viewInfo.DefaultView,
                            Hidden: viewInfo.Hidden,
                            JSLink: viewInfo.JSLink,
                            MobileDefaultView: viewInfo.MobileDefaultView,
                            MobileView: viewInfo.MobileView,
                            RowLimit: viewInfo.RowLimit,
                            Tabular: viewInfo.TabularView,
                            ViewName: viewInfo.Title,
                            ViewFields: viewInfo.ViewFields.Items.results,
                            ViewQuery: viewInfo.ViewQuery
                        });
                    }

                    // Parse the calculated fields
                    for (let i = 0; i < calcFields.length; i++) {
                        let calcField = calcFields[i];

                        if (fields[calcField.InternalName] == null) {
                            let parser = new DOMParser();
                            let schemaXml = parser.parseFromString(calcField.SchemaXml, "application/xml");

                            // Get the formula
                            let formula = schemaXml.querySelector("Formula");

                            // Parse the field refs
                            let fieldRefs = schemaXml.querySelectorAll("FieldRef");
                            for (let j = 0; j < fieldRefs.length; j++) {
                                let fieldRef = fieldRefs[j].getAttribute("Name");

                                // Ensure the field exists
                                let field = list.getField(fieldRef);
                                if (field) {
                                    // Calculated formulas are supposed to contain the display name
                                    // Replace any instance of the internal field w/ the correct format
                                    let regexp = new RegExp(fieldRef, "g");
                                    formula.innerHTML = formula.innerHTML.replace(regexp, "[" + field.Title + "]");
                                }
                            }

                            // Append the field
                            fields[calcField.InternalName] = true;
                            cfgProps.ListCfg[0].CustomFields.push({
                                name: calcField.InternalName,
                                schemaXml: schemaXml.querySelector("Field").outerHTML
                            });
                        }
                    }

                    // Resolve the request
                    resolve({
                        cfg: cfgProps,
                        lookupFields
                    });
                }
            });
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
                                let formValues = this._form.getValues();

                                // Set the list name
                                let listData = formValues["SourceList"].data as Types.SP.List;

                                // Copy the list
                                this.createListConfiguration(appItem, formValues["WebUrl"], listData);
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
            // Get the list configurations and append/replace it
            let listCfg = appItem.ListConfigurations;

            // Converting string to object may fail
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
            appItem.update({ ListConfigurations: listCfg }).execute(resolve);
        });
    }
}