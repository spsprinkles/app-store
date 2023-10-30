import { InstallationRequired, LoadingDialog } from "dattatable";
import { Components } from "gd-sprest-bs";
import { Configuration } from "./cfg";
import { Security } from "./security";
import Strings from "./strings";

/**
 * Installation Modal
 */
export class InstallationModal {
    // Shows the modal
    static show(showFl: boolean = false) {
        // See if an installation is required
        InstallationRequired.requiresInstall({ cfg: Configuration }).then(installFl => {
            let customErrors: Components.IListGroupItem[] = [];

            // See if the security groups exist
            let securityGroupsExist = true;
            if (Security.DeveloperGroup == null || Security.ManagerGroup == null) {
                // Set the flag
                securityGroupsExist = false;

                // Add an error
                customErrors.push({
                    content: "Security groups are not installed.",
                    type: Components.ListGroupItemTypes.Danger
                });
            }

            // See if an install is required
            if (installFl || customErrors.length > 0 || showFl) {
                // Show the dialog
                InstallationRequired.showDialog({
                    errors: customErrors,
                    onFooterRendered: el => {
                        // See if the security group doesn't exist
                        if (!securityGroupsExist || showFl) {
                            // Add the custom install button
                            Components.Tooltip({
                                el,
                                content: "Creates the security groups.",
                                type: Components.ButtonTypes.OutlinePrimary,
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
    }
}