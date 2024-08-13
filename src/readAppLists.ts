import { ListConfig, LoadingDialog } from "dattatable";
import { Components, Helper, SPTypes, Types, Web } from "gd-sprest-bs";
import { DataSource, IAppStoreItem } from "./ds";
import { getListTemplateUrl } from "./strings";
import { CreateAppLists } from "./createAppLists";

/**
 * Read App Lists
 * Reads an existing solution's list and generates a configuration for it.
 */
export class ReadAppLists {
    private static _elListCfg: HTMLElement = null;
    private static _form: Components.IForm = null;

    // Method to create the list configuration
    private static createListConfiguration(appItem: IAppStoreItem, srcWebUrl: string, srcList: Types.SP.List): PromiseLike<void> {
        // Show a loading dialog
        LoadingDialog.setHeader("Copying the List");
        LoadingDialog.setBody("Initializing the request...");
        LoadingDialog.show();

        // Return a promise
        return new Promise((resolve, reject) => {
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
                    ListConfig.generate({
                        showDialog: true,
                        srcWebUrl,
                        srcList
                    }).then(srcListCfg => {
                        // Save a copy of the configuration
                        let strConfig = JSON.stringify(srcListCfg.cfg);

                        // Validate the lookup fields
                        ListConfig.validateLookups({
                            cfg: srcListCfg.cfg,
                            dstUrl: web.ServerRelativeUrl,
                            lookupFields: srcListCfg.lookupFields,
                            showDialog: true,
                            srcList,
                            srcWebUrl,
                        }).then((listCfg) => {
                            // Test the configuration
                            CreateAppLists.installConfiguration(listCfg, web.ServerRelativeUrl).then(lists => {
                                // Update the list configuration
                                this.updateListConfiguration(appItem, JSON.parse(strConfig)).then(() => {
                                    // Hide the loading dialog
                                    LoadingDialog.hide();

                                    // Show the results
                                    CreateAppLists.showResults(listCfg, web.ServerRelativeUrl, lists, true);

                                    // Resolve the request
                                    resolve();
                                }, reject);
                            }, reject);
                        }, reject);
                    }, reject);
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

    // Renders the footer
    static renderFooter(el: HTMLElement, appItem: IAppStoreItem, webUrl?: string) {
        // See if the footer exists
        let elFooter = el.querySelector("#footer");
        if (elFooter) {
            // Remove the element
            elFooter.parentElement.removeChild(elFooter);
        }

        // Set the footer
        Components.TooltipGroup({
            el,
            id: "footer",
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
                                    // Refresh the list configuration
                                    this.renderListConfiguration(item, webUrl);

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
                                        this.renderListConfiguration(item, webUrl);

                                        // Refresh the footer
                                        this.renderFooter(el, item, webUrl);
                                    });
                                }, err => {
                                    // Log the error
                                    console.log("Error creating the configuration", err);

                                    // Show the error if it's a string
                                    let ctrl = this._form.getControl("SourceList");
                                    if (typeof (err) === "string") {
                                        // Update the validation
                                        ctrl.updateValidation(ctrl.el, {
                                            isValid: false,
                                            invalidMessage: err
                                        });
                                    } else {
                                        // Update the validation
                                        ctrl.updateValidation(ctrl.el, {
                                            isValid: false,
                                            invalidMessage: "Error creating the configuration. Check the console log for details."
                                        });
                                    }

                                    // Hide the loading dialog
                                    LoadingDialog.hide();
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

        // Render the existing list configuration
        this._elListCfg = document.createElement("div");
        this._elListCfg.id = "list-cfg";
        el.appendChild(this._elListCfg);
        this.renderListConfiguration(appItem, webUrl);

        // Set the body
        let label = document.createElement("label");
        label.classList.add("my-2");
        el.appendChild(label);

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

    private static renderListConfiguration(appItem: IAppStoreItem, webUrl?: string, newAppConfig?: Helper.ISPConfigProps) {
        // Clear the element
        while (this._elListCfg.firstChild) { this._elListCfg.removeChild(this._elListCfg.firstChild); }

        // Render the current lists in the configuration
        try {
            // Get the configuration
            let appConfig: Helper.ISPConfigProps = newAppConfig || JSON.parse(appItem.ListConfigurations);
            if (appConfig && appConfig.ListCfg.length > 0) {
                let items: Components.IListGroupItem[] = [];

                // Parse the list configurations
                for (let i = 0; i < appConfig.ListCfg.length; i++) {
                    // Add the item
                    items.push({
                        data: i,
                        className: "d-flex justify-content-between",
                        content: appConfig.ListCfg[i].ListInformation.Title,
                        onRender: (el, item) => {
                            // Add the badges
                            let elBadges = document.createElement("div");
                            el.appendChild(elBadges);

                            // Append a badge to remove the item
                            Components.Badge({
                                el: elBadges,
                                content: "Delete",
                                type: Components.BadgeTypes.Danger,
                                onClick: () => {
                                    // Remove the item
                                    appConfig.ListCfg.splice(i, 1);

                                    // Render the list configuration
                                    this.renderListConfiguration(appItem, webUrl, appConfig);
                                }
                            });

                            // See if this is not the first item
                            if (i > 0) {
                                // Add a badge to move this item up
                                Components.Badge({
                                    el: elBadges,
                                    className: "ms-2",
                                    content: "Move Up",
                                    onClick: () => {
                                        // Get the item
                                        let item = appConfig.ListCfg.splice(i, 1)[0];

                                        // Move the item up one spot in the array
                                        appConfig.ListCfg.splice(i - 1, 0, item);

                                        // Render the list configuration
                                        this.renderListConfiguration(appItem, webUrl, appConfig);
                                    }
                                });
                            }
                        }
                    });
                }

                // Add a label for the list
                this._elListCfg.innerHTML = `<label class="my-2">
                            Below are the current lists associated with this solution.
                            ${items.length > 1 ? "Order matters when using lookup fields. Use the arrows to reorder the lists accordingly." : ""}
                        </label>`;

                // Render the list
                Components.ListGroup({
                    el: this._elListCfg,
                    items
                });

                // Add a save button
                let elSaveButton = document.createElement("div");
                newAppConfig ? null : elSaveButton.classList.add("d-none");
                elSaveButton.classList.add("mt-2");
                elSaveButton.classList.add("d-flex");
                elSaveButton.classList.add("justify-content-end");
                this._elListCfg.appendChild(elSaveButton);
                Components.Button({
                    el: elSaveButton,
                    isSmall: true,
                    text: "Update Configuration",
                    onClick: () => {
                        // Update the item
                        appItem.update({ ListConfigurations: JSON.stringify(appConfig) }).execute(() => {
                            // Update the item
                            DataSource.refresh(appItem.Id).then((item: IAppStoreItem) => {
                                // Refresh the form
                                this.renderListConfiguration(item, webUrl);

                                // Refresh the footer
                                this.renderFooter(this._elListCfg.parentElement, item, webUrl);
                            });
                        });
                    }
                });
            }
        } catch { }
    }

    // Updates the list configuration for the item
    private static updateListConfiguration(appItem: IAppStoreItem, cfg: Helper.ISPConfigProps): PromiseLike<void> {
        // Return a promise
        return new Promise(resolve => {
            // Get the list configurations and append/replace it
            let listConfigs: string = null;

            // Converting string to object may fail
            try {
                // Get the app configuration
                let appConfig: Helper.ISPConfigProps = JSON.parse(appItem.ListConfigurations);

                // Parse the configurations
                Helper.Executor(cfg.ListCfg, listCfg => {
                    let listName = listCfg.ListInformation.Title;

                    // Parse the lists
                    let foundFl = false;
                    for (let i = 0; i < appConfig.ListCfg.length; i++) {
                        // See if this is the target list
                        if (appConfig.ListCfg[i].ListInformation.Title == listName) {
                            // Replace it
                            appConfig.ListCfg[i] = listCfg;
                            foundFl = true;
                            break;
                        }
                    }

                    // See if it wasn't found
                    if (!foundFl) {
                        // Append the configuration
                        appConfig.ListCfg.push(listCfg);
                    }
                }).then(() => {
                    // Set the configuration
                    listConfigs = JSON.stringify(appConfig);
                });
            } catch {
                // Set the configuration
                listConfigs = JSON.stringify(cfg);
            }

            // Update the item
            appItem.update({ ListConfigurations: listConfigs }).execute(resolve);
        });
    }
}