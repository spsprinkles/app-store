import { List, LoadingDialog, Modal } from "dattatable";
import { Components, ContextInfo, Helper, SPTypes, Types, Web } from "gd-sprest-bs";
import { IAppStoreItem } from "../ds";
import { getListTemplateUrl } from "../strings";

/**
 * Create Templates
 */
export class CreateTemplates {
    private static _form: Components.IForm = null;

    // Method to copy the list
    private static copyList(appTitle: string, srcListName: string, dstWebUrl: string, dstListName: string): PromiseLike<string> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Getting the source list information
            LoadingDialog.setBody("Loading source list...");

            // Get the list information
            let srcWebUrl = getListTemplateUrl();
            var list = new List({
                listName: srcListName,
                webUrl: srcWebUrl,
                itemQuery: { Filter: "Id eq 0" },
                onInitError: () => {
                    // Reject the request
                    reject("Error loading the list information. Please check your permission to the source list.");
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
                    cfg.setWebUrl(dstWebUrl);
                    cfg.install().then(() => {
                        // Update the loading dialog
                        LoadingDialog.setBody("Validating the lookup field(s)...");

                        // Parse the lookup fields
                        Helper.Executor(lookupFields, lookupField => {
                            // Return a promise
                            return new Promise(resolve => {
                                // Get the lookup field source list
                                Web(srcWebUrl).Lists().getById(lookupField.LookupList).execute(srcLookupList => {
                                    // The lookup list template format will need to remove the app title for the destination list.
                                    let dstLookupList = srcLookupList.Title.replace(appTitle, "").trim();

                                    // Get the lookup list in the destination site
                                    Web(dstWebUrl).Lists(dstLookupList).execute(dstList => {
                                        // Get the context for the destination web
                                        ContextInfo.getWeb(dstWebUrl).execute(contextInfo => {
                                            // Update the field schema xml
                                            let fieldDef = lookupField.SchemaXml.replace(`List="${lookupField.LookupList}"`, `List="{${dstList.Id}}"`);
                                            Web(dstWebUrl, { requestDigest: contextInfo.GetContextWebInformation.FormDigestValue })
                                                .Lists(dstListName).Fields(lookupField.InternalName).update({
                                                    SchemaXml: fieldDef
                                                }).execute(() => {
                                                    // Updated the lookup list
                                                    console.log(`Updated the lookup field '${lookupField.InternalName}' in lookup list successfully.`);

                                                    // Check the next field
                                                    resolve(null);
                                                });
                                        });
                                    }, () => {
                                        // Error getting the lookup list
                                        console.error(`Error getting the lookup list '${dstLookupList}' from web '${dstWebUrl}'.`);

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
                            resolve(dstListName);
                        });
                    });
                }
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

    // Renders the footer
    static renderFooter(el: HTMLElement, appItem: IAppStoreItem, listNames: string[]) {
        // Clear the footer
        while (el.firstChild) { el.removeChild(el.firstChild); }

        // Set the footer
        Components.TooltipGroup({
            el,
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
                                        // Show a loading dialog
                                        LoadingDialog.setHeader("Copying the List");
                                        LoadingDialog.setBody("Initializing the request...");
                                        LoadingDialog.show();

                                        // Parse the lists to copy
                                        Helper.Executor(listNames, listName => {
                                            // Return a promise
                                            return new Promise(resolve => {
                                                // Set the destination list name
                                                let dstListName = listName.replace(appItem.Title, "").trim();

                                                // Copy the list
                                                this.copyList(appItem.Title, listName, dstWebUrl, dstListName).then(
                                                    () => {
                                                        // Copy the next list
                                                        resolve(null);
                                                    },
                                                    err => {
                                                        // Show the error message
                                                        ctrlWeb.updateValidation(ctrlWeb.el, {
                                                            isValid: false,
                                                            invalidMessage: err
                                                        });
                                                    }
                                                );
                                            });
                                        }).then(() => {
                                            // Hide the dialog
                                            LoadingDialog.hide();

                                            // Update the validation
                                            ctrlWeb.updateValidation(ctrlWeb.el, {
                                                isValid: true,
                                                validMessage: "List copied successfully"
                                            });
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
                },
                {
                    content: "Close this dialog",
                    btnProps: {
                        text: "Close",
                        type: Components.ButtonTypes.OutlineSecondary,
                        onClick: () => {
                            // Close the dialog
                            Modal.hide();
                        }
                    }
                }
            ]
        });
    }

    // Renders the main form
    static renderForm(el: HTMLElement, appItem: IAppStoreItem, listNames: string[]) {
        // Get the list templates associated w/ this item
        if (listNames.length > 0) {
            // Set the body
            el.innerHTML = `<label class="my-2">This will create the list(s) required for this app in the web specified.</label>`;

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
}