import { List, LoadingDialog, Modal } from "dattatable";
import { Components, Helper, SPTypes, Web } from "gd-sprest-bs";
import * as Common from "./common";
import { IAppStoreItem } from "./ds";
import { ReadAppLists } from "./readAppLists";

/**
 * Create App Lists
 */
export class CreateAppLists {
    private static _form: Components.IForm = null;

    // Method to create the lists
    private static createLists(cfgProps: Helper.ISPConfigProps, webUrl: string): PromiseLike<List[]> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Create the configuration
            let cfg = Helper.SPConfig(cfgProps);

            // Update the loading dialog
            LoadingDialog.setBody("Deleting the existing lists...");

            // Uninstall the solution
            cfg.uninstall().then(() => {
                // Install the solution
                cfg.install().then(() => {
                    let lists: List[] = [];

                    // Parse the lists
                    Helper.Executor(cfgProps.ListCfg, listCfg => {
                        // Return a promise
                        return new Promise(resolve => {
                            // Test the list
                            this.testList(listCfg, webUrl).then(list => {
                                // Append the list
                                lists.push(list);

                                // Check the next list
                                resolve(null);
                            });
                        });
                    }, reject).then(() => {
                        // Resolve the request
                        resolve(lists);
                    });
                }, reject);
            });
        });
    }

    // Fixes the lookup fields
    private static testList(listCfg: Helper.ISPCfgListInfo, webUrl: string): PromiseLike<List> {
        // Update the loading dialog
        LoadingDialog.setBody("Validating the lookup field(s)...");

        // Return a promise
        return new Promise((resolve, reject) => {
            // Create the list
            let list = new List({
                listName: listCfg.ListInformation.Title,
                webUrl,
                onInitialized: () => {
                    // Resolve the list
                    resolve(list);
                },
                onInitError: reject
            });
        });
    }

    // Checks if the user has access to the destination web
    private static hasAccess(webUrl): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve, reject) => {
            Web(webUrl).query({ Expand: ["EffectiveBasePermissions"] }).execute(
                // Exists
                web => {
                    // Ensure the user doesn't have permission to manage lists
                    if (!Helper.hasPermissions(web.EffectiveBasePermissions, [SPTypes.BasePermissionTypes.ManageLists])) {
                        // Reject the request
                        reject("You do not have permission to copy templates on this web.");
                    } else {
                        // Resolve the request
                        resolve()
                    }
                },
                // Doesn't exist
                () => {
                    // Reject the request
                    reject("The site doesn't exist or you do not have access to it.");
                }
            );
        });
    }

    // Installs the configuration
    static installConfiguration(cfg: Helper.ISPConfigProps, webUrl: string): PromiseLike<List[]> {
        // Show a loading dialog
        LoadingDialog.setHeader("Creating the List");
        LoadingDialog.setBody("Initializing the request...");
        LoadingDialog.show();

        // Return a promise
        return new Promise(resolve => {
            // Create the list(s)
            this.createLists(cfg, webUrl).then(lists => {
                // Hide the dialog
                LoadingDialog.hide();

                // Resolve the request
                resolve(lists);
            });
        });
    }

    // Renders the main form
    static renderForm(el: HTMLElement, appItem: IAppStoreItem, cfgProps: Helper.ISPConfigProps) {
        // Get the list templates associated w/ this item
        if (cfgProps && cfgProps.ListCfg.length > 0) {
            // Set the body
            el.innerHTML = `<label class="mb-2">This will create the list(s) required for this app in the web specified.</label>`;

            // Render a form
            this._form = Components.Form({
                el,
                controls: [
                    {
                        name: "WebUrl",
                        title: "Destination Web Url",
                        type: Components.FormControlTypes.TextField,
                        description: "The destination web to copy the list(s) to",
                        required: true,
                        errorMessage: "A relative web url is required (Ex. /sites/dev)"
                    }
                ]
            });
        } else {
            // Set the body
            el.innerHTML = `<p>No list templates exist for this app: ${appItem.Title}.</p>`;
        }
    }

    // Renders the modal
    static renderModal(appItem: IAppStoreItem, listCfg: string) {
        // Clear the modal
        Modal.clear();

        // Set the header
        Modal.setHeader("Deploy Dataset: " + appItem.Title);
        Modal.HeaderElement.prepend(Common.getIcon(28, 28, appItem.AppType + (appItem.AppType.startsWith('Power ') ? ' Template' : ''), 'icon-svg me-2'));

        let cfgProps: Helper.ISPConfigProps = null;
        try {
            // Get the configurations
            cfgProps = JSON.parse(listCfg);
        } catch { }

        // Render the form
        this.renderForm(Modal.BodyElement, appItem, cfgProps);

        // Set the footer
        Components.TooltipGroup({
            el: Modal.FooterElement,
            tooltips: [
                {
                    content: "Copy the associated lists to the destination web",
                    btnProps: {
                        text: "Start Copy",
                        type: Components.ButtonTypes.OutlinePrimary,
                        onClick: () => {
                            // Validate the form
                            if (this._form.isValid()) {
                                let ctrlWeb = this._form.getControl("WebUrl");
                                let dstWebUrl = ctrlWeb.getValue();

                                // Ensure the user has access to the destination web
                                this.hasAccess(dstWebUrl).then(
                                    // Has access to copy templates
                                    () => {
                                        // Install the configuration
                                        this.installConfiguration(cfgProps, dstWebUrl).then(lists => {
                                            // Show the results
                                            this.showResults(cfgProps, lists);
                                        });
                                    },

                                    // Error accessing the web
                                    err => {
                                        // Hide the dialog
                                        LoadingDialog.hide();

                                        // Set the validation
                                        ctrlWeb.updateValidation(ctrlWeb.el, {
                                            invalidMessage: err,
                                            isValid: false
                                        });
                                    }
                                );
                            }
                        }
                    }
                }
            ]
        });

        // Show the modal
        Modal.show();
    }

    // Shows the results of the copy
    static showResults(cfg: Helper.ISPConfigProps, lists: List[], showDeleteFl: boolean = false) {
        // Clear the modal
        Modal.clear();

        // Set the header
        Modal.setHeader("Create Lists");

        // Set the body
        Modal.setBody("Click on the link(s) below to access the list settings for validation.");

        // Parse the lists
        let items: Components.IListGroupItem[] = null;
        for (let i = 0; i < lists.length; i++) {
            let list = lists[i];

            // Add the list links
            items.push({
                data: list,
                content: Components.ButtonGroup({
                    buttonType: Components.ButtonTypes.OutlinePrimary,
                    isSmall: true,
                    buttons: [
                        {
                            text: "List",
                            onClick: () => {
                                // Go to the list
                                window.open(list.ListUrl, "_blank");
                            }
                        },
                        {
                            className: "ms-2",
                            text: "List Settings",
                            onClick: () => {
                                // Go to the list
                                window.open(list.ListSettingsUrl, "_blank");
                            }
                        }
                    ]
                }).el
            })
        }

        Components.ListGroup({
            el: Modal.BodyElement,
            items
        });

        // See if we are deleting the list
        if (showDeleteFl) {
            // Render a delete button
            Components.Tooltip({
                el: Modal.FooterElement,
                content: "Click to delete the test lists.",
                btnProps: {
                    text: "Delete Lists",
                    type: Components.ButtonTypes.OutlineDanger,
                    onClick: () => {
                        // Show a loading dialog
                        LoadingDialog.setHeader("Deleteing List(s)");
                        LoadingDialog.setBody("This will close after the lists are removed...");
                        LoadingDialog.show();

                        // Uninstall the configuration
                        Helper.SPConfig(cfg).uninstall().then(() => {
                            // Hide the loading dialog
                            LoadingDialog.hide();
                        });
                    }
                }
            });
        }

        // Show the modal
        Modal.show();
    }
}