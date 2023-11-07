import { InstallationRequired, LoadingDialog } from "dattatable";
import { Components, Web } from "gd-sprest-bs";
import { Configuration } from "./cfg";
import { Security } from "./security";
import Strings, { getListTemplateUrl } from "./strings";

/**
 * Installation Modal
 */
export class InstallationModal {
    // Creates the sub-web
    private static createSubWeb(): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Show a loading dialog
            LoadingDialog.setHeader("Creating Sub-Web");
            LoadingDialog.setBody("This will close after the web is created...");
            LoadingDialog.show();

            // Create the web
            Web().WebInfos().add(Strings.ListTemplateWebInfo).execute(
                // Created
                web => {
                    // Hide the loading dialog
                    LoadingDialog.hide();

                    // Resolve the request
                    resolve();
                },

                // Not created
                () => {
                    // Hide the loading dialog
                    LoadingDialog.hide();

                    // Log
                    console.error("[" + Strings.ProjectName + "] Error creating the list templates sub-web.");

                    // Reject the request
                    reject();
                }
            )
        });
    }

    // Checks to see if the sub-web exists
    private static hasSubWeb(): PromiseLike<boolean> {
        // Return a promise
        return new Promise((resolve) => {
            // See if the web exists
            Web(getListTemplateUrl()).execute(
                // Exists
                () => { resolve(true); },

                // Doesn't exist
                () => { resolve(false); }
            );
        });
    }

    // Shows the modal
    static show(showFl: boolean = false) {
        // See if an installation is required
        InstallationRequired.requiresInstall({ cfg: Configuration }).then(installFl => {
            // See if the sub-web exists
            this.hasSubWeb().then(hasSubWeb => {
                let customErrors: Components.IListGroupItem[] = [];

                // See if the sub-web exists
                if (!hasSubWeb) {
                    // Add an error
                    customErrors.push({
                        content: "Sub-Web for list templates does not exist",
                        type: Components.ListGroupItemTypes.Danger
                    });
                }

                // See if the security groups exist
                let securityGroupsExist = true;
                if (Security.DeveloperGroup == null || Security.ManagerGroup == null) {
                    // Set the flag
                    securityGroupsExist = false;

                    // Add an error
                    customErrors.push({
                        content: "Security groups are not installed",
                        type: Components.ListGroupItemTypes.Danger
                    });
                }

                // See if an install is required
                if (installFl || customErrors.length > 0 || showFl) {
                    // Show the dialog
                    InstallationRequired.showDialog({
                        errors: customErrors,
                        onFooterRendered: el => {
                            // See if the sub-web doesn't exist
                            if (!hasSubWeb) {
                                // Add the custom install button
                                Components.Tooltip({
                                    el,
                                    content: "Create the sub-web for storing list templates",
                                    type: Components.ButtonTypes.OutlinePrimary,
                                    options: {
                                        theme: "sharepoint"
                                    },
                                    btnProps: {
                                        text: "Create Sub-Web",
                                        onClick: () => {
                                            // Create the sub-web
                                            this.createSubWeb().then(() => {
                                                // Refresh the page
                                                window.location.reload();
                                            });
                                        }
                                    }
                                })
                            }

                            // See if the security group doesn't exist
                            if (!securityGroupsExist || showFl) {
                                // Add the custom install button
                                Components.Tooltip({
                                    el,
                                    content: "Create the security groups",
                                    type: Components.ButtonTypes.OutlinePrimary,
                                    options: {
                                        theme: "sharepoint"
                                    },
                                    btnProps: {
                                        text: "Security",
                                        isDisabled: !InstallationRequired.ListsExist,
                                        onClick: () => {
                                            // Create the security groups
                                            Security.show(() => {
                                                // Refresh the page
                                                window.location.reload();
                                            });
                                        }
                                    }
                                });
                            }
                        }
                    });
                } else {
                    // Log
                    console.error("[" + Strings.ProjectName + "] Error initializing the solution.");
                }
            });
        });
    }
}