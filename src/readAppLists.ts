import { List, LoadingDialog } from "dattatable";
import { Components, Helper, SPTypes, Types, Web } from "gd-sprest-bs";
import { DataSource, IAppStoreItem } from "./ds";
import { getListTemplateUrl } from "./strings";
import { CreateAppLists } from "./createAppLists";

/**
 * Read App Lists
 * Reads an existing solution's list and generates a configuration for it.
 */
export class ReadAppLists {
    private static _form: Components.IForm = null;

    // Method to create the list configuration
    private static createListConfiguration(appItem: IAppStoreItem, srcWebUrl: string, srcList: Types.SP.List): PromiseLike<void> {
        // Show a loading dialog
        LoadingDialog.setHeader("Copying the List");
        LoadingDialog.setBody("Initializing the request...");
        LoadingDialog.show();

        // Return a promise
        return new Promise(resolve => {
            // Ensure the user has the correct permissions to create the list
            let dstWebUrl = getListTemplateUrl();
            Web(dstWebUrl).query({ Expand: ["EffectiveBasePermissions"] }).execute(
                // Exists
                web => {
                    // Ensure the user doesn't have permission to manage lists
                    if (!Helper.hasPermissions(web.EffectiveBasePermissions, [SPTypes.BasePermissionTypes.ManageLists])) {
                        // Resolve the request
                        console.error("You do not have permission to copy templates on this web.");
                        LoadingDialog.hide();
                        resolve();
                        return;
                    }

                    // Generate the list configuration
                    this.generateListConfiguration(srcWebUrl, srcList).then(listCfg => {
                        // Save a copy of the configuration
                        let strConfig = JSON.stringify(listCfg.cfg);

                        // Test the configuration
                        CreateAppLists.installConfiguration(listCfg.cfg, web.ServerRelativeUrl).then(lists => {
                            // Update the list configuration
                            this.updateListConfiguration(appItem, srcList.Title, JSON.parse(strConfig)).then(() => {
                                // Hide the loading dialog
                                LoadingDialog.hide();

                                // Show the results
                                CreateAppLists.showResults(listCfg.cfg, web.ServerRelativeUrl, lists, true);

                                // Resolve the request
                                resolve();
                            });
                        });
                    });
                },

                // Doesn't exist
                () => {
                    // Hide the loading dialog
                    LoadingDialog.hide();

                    // Resolve the request
                    console.error("Error getting the target web. You may not have permissions or it doesn't exist.");
                    resolve();
                }
            );
        });
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
                            let fldInfo: Types.SP.Field = list.getField(ct.FieldLinks.results[j].Name);

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
                            }
                            // Else, see if this is a lookup field
                            else if (fldInfo.FieldTypeKind == SPTypes.FieldType.Lookup) {
                                // Add the field
                                lookupFields.push(fldInfo);
                            }
                            // Else, ensure the field hasn't been added
                            else if (fields[fldInfo.InternalName] == null) {
                                // Add the field information
                                fields[fldInfo.InternalName] = true;
                                cfgProps.ListCfg[0].CustomFields.push({
                                    name: fldInfo.InternalName,
                                    schemaXml: fldInfo.SchemaXml
                                });
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
                                }
                                // Else, see if this is a lookup field
                                else if (field.FieldTypeKind == SPTypes.FieldType.Lookup) {
                                    // Add the field
                                    lookupFields.push(field);
                                } else {
                                    // Append the field
                                    fields[field.InternalName] = true;
                                    cfgProps.ListCfg[0].CustomFields.push({
                                        name: field.InternalName,
                                        schemaXml: field.SchemaXml
                                    });
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

                    // Update the loading dialog
                    LoadingDialog.setBody("Analyzing the lookup fields...");

                    // Parse the lookup fields
                    Helper.Executor(lookupFields, lookupField => {
                        // Skip the field, if it was already added
                        if (fields[lookupField.InternalName]) { return; }

                        // Return a promise
                        return new Promise((resolve) => {
                            // Get the lookup list
                            Web(srcWebUrl).Lists().getById(lookupField.LookupList).execute(
                                list => {
                                    // Add the lookup list field
                                    fields[lookupField.InternalName] = true;
                                    cfgProps.ListCfg[0].CustomFields.push({
                                        description: lookupField.Description,
                                        fieldRef: lookupField.PrimaryFieldId,
                                        hidden: lookupField.Hidden,
                                        id: lookupField.Id,
                                        indexed: lookupField.Indexed,
                                        listName: list.Title,
                                        multi: lookupField.AllowMultipleValues,
                                        name: lookupField.InternalName,
                                        readOnly: lookupField.ReadOnlyField,
                                        relationshipBehavior: lookupField.RelationshipDeleteBehavior,
                                        required: lookupField.Required,
                                        showField: lookupField.LookupField,
                                        title: lookupField.Title,
                                        type: Helper.SPCfgFieldType.Lookup
                                    } as Helper.IFieldInfoLookup);

                                    // Check the next field
                                    resolve(null);
                                },

                                err => {
                                    // Broken lookup field, don't add it
                                    console.log("Error trying to find lookup list for field '" + lookupField.InternalName + "' with id: " + lookupField.LookupList);
                                    resolve(null);
                                }
                            )
                        });
                    }).then(() => {
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
                    });
                }
            });
        });
    }

    // Renders the footer
    static renderFooter(el: HTMLElement, appItem: IAppStoreItem, webUrl?: string) {
        // Set the footer
        Components.TooltipGroup({
            el,
            className: "float-end",
            tooltips: [
                {
                    content: "Clears the list templates for this item.",
                    btnProps: {
                        text: "Clear Configuration",
                        type: Components.ButtonTypes.OutlinePrimary,
                        isDisabled: appItem.ListConfigurations ? false : true,
                        onClick: () => {
                            // Show a loading dialog
                            LoadingDialog.setHeader("Clearing Configuration");
                            LoadingDialog.setBody("The list configurations are being removed...");
                            LoadingDialog.show();

                            // Clear the value in the item
                            appItem.update({
                                ListConfigurations: null
                            }).execute(() => {
                                // Refresh the item
                                DataSource.refresh(appItem.Id).then((item: IAppStoreItem) => {
                                    // Refresh the form
                                    this.renderForm(el, item, webUrl);

                                    // Refresh the footer
                                    this.renderFooter(el, item, webUrl);

                                    // Hide the loading dialog
                                    LoadingDialog.hide();
                                });
                            })
                        }
                    }
                },
                {
                    content: "Tests the configuration by creating them in a sub-web.",
                    btnProps: {
                        text: "Test Configuration",
                        type: Components.ButtonTypes.OutlinePrimary,
                        isDisabled: appItem.ListConfigurations ? false : true,
                        onClick: () => {
                            let dstWebUrl = getListTemplateUrl();

                            // Show a loading dialog
                            LoadingDialog.setHeader("Testing Configuration");
                            LoadingDialog.setBody("The lists are being created...");
                            LoadingDialog.show();

                            // Get the configuration
                            let cfg = JSON.parse(appItem.ListConfigurations);

                            // Test the configuration
                            CreateAppLists.installConfiguration(cfg, dstWebUrl).then(lists => {
                                // Hide the loading dialog
                                LoadingDialog.hide();

                                // Show the results
                                CreateAppLists.showResults(cfg, dstWebUrl, lists, true);
                            })
                        }
                    }
                },
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
                                let webUrl: string = formValues["WebUrl"];
                                let listData = formValues["SourceList"].data as Types.SP.List;

                                // Copy the list
                                this.createListConfiguration(appItem, webUrl, listData).then(() => {
                                    // Refresh the item
                                    DataSource.refresh(appItem.Id).then((item: IAppStoreItem) => {
                                        // Refresh the form
                                        this.renderForm(el, item, webUrl);

                                        // Refresh the footer
                                        this.renderFooter(el, item, webUrl);
                                    });
                                });
                            }
                        }
                    }
                }
            ]
        });
    }

    // Renders the main form
    static renderForm(el: HTMLElement, appItem: IAppStoreItem, webUrl?: string) {
        // Clear the form
        while (el.firstChild) { el.removeChild(el.firstChild); }

        // Render the current lists in the configuration
        try {
            // Get the configuration
            let appConfig: Helper.ISPConfigProps = JSON.parse(appItem.ListConfigurations);
            if (appConfig && appConfig.ListCfg.length > 0) {
                let items: Components.IListGroupItem[] = [];

                // Parse the list configurations
                for (let i = 0; i < appConfig.ListCfg.length; i++) {
                    // Add the item
                    items.push({
                        content: appConfig.ListCfg[i].ListInformation.Title,
                        badge: i == 0 ? null : {
                            content: "Move Up",
                            onClick: () => {

                            }
                        }
                    });
                }

                // Add a label for the list
                el.innerHTML = `<label class="my-2">
                    Below are the current lists associated with this solution.
                    ${items.length > 1 ? "Order matters when using lookup fields. Use the arrows to reorder the lists accordingly." : ""}
                </label>`;

                // Render the list
                Components.ListGroup({
                    el,
                    items
                });
            }
        } catch { }

        // Set the body
        el.innerHTML += `<label class="my-2">Use this form to add or update a list template for this app.</label>`;

        // Render a form
        this._form = Components.Form({
            el,
            className: "mt-3",
            controls: [
                {
                    name: "WebUrl",
                    label: "Source Web Url:",
                    type: Components.FormControlTypes.TextField,
                    description: "The source web url containing the lists to import.",
                    required: true,
                    errorMessage: "A relative web url is required. (Ex. /sites/dev)",
                    value: webUrl
                },
                {
                    name: "SourceList",
                    label: "Select a List:",
                    type: Components.FormControlTypes.Dropdown,
                    description: "Select a list to import.",
                    required: true,
                    errorMessage: "A list is required."
                }
            ]
        });
    }

    // Updates the list configuration for the item
    private static updateListConfiguration(appItem: IAppStoreItem, listName: string, cfg: Helper.ISPConfigProps): PromiseLike<void> {
        // Return a promise
        return new Promise(resolve => {
            // Get the list configurations and append/replace it
            let listCfg: string = null;

            // Converting string to object may fail
            try {
                // Get the app configuration
                let appConfig: Helper.ISPConfigProps = JSON.parse(appItem.ListConfigurations);

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