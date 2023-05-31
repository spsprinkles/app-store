import { LoadingDialog } from "dattatable";
import { Components, Helper } from "gd-sprest-bs";
import { DataSource, IAppStoreItem } from "./ds";

/**
 * Forms
 */
export class Forms {
    // Configures the form
    private static configureForm(props: Components.IListFormEditProps): Components.IListFormEditProps {
        // Set the control rendering event
        props.onControlRendering = (ctrl, fld) => {
            // See if this is a url field
            if (fld.InternalName == "Icon" || fld.InternalName.indexOf("ScreenShot") == 0) {
                // Make the field read-only
                ctrl.isReadonly = true;
            }
            // Else, see if this is a link
            else if (fld.InternalName == "AdditionalInformation") {
                // Update the props
                (ctrl as Components.IFormControlUrlProps).showDescription = false;
            }
        }

        // Set the control rendered event
        props.onControlRendered = (ctrl, fld) => {
            // See if this is a url field
            if (fld.InternalName == "Icon" || fld.InternalName.indexOf("ScreenShot") == 0) {
                // Set a click event
                ctrl.textbox.elTextbox.addEventListener("click", () => {
                    // Display a file upload dialog
                    Helper.ListForm.showFileDialog().then(file => {
                        // Clear the value
                        ctrl.textbox.setValue("");

                        // Get the file name
                        let fileName = file.name.toLowerCase();

                        // Validate the file type
                        if (fileName.endsWith(".png") || fileName.endsWith(".jpg") || fileName.endsWith(".jpeg") || fileName.endsWith(".gif")) {
                            // Show a loading dialog
                            LoadingDialog.setHeader("Reading the File");
                            LoadingDialog.setBody("This will close after the file is converted...");
                            LoadingDialog.show();

                            // Convert the file
                            let reader = new FileReader();
                            reader.onloadend = () => {
                                let value = reader.result as string;

                                // Set the value
                                ctrl.textbox.setValue(value.replace("data:", '').replace(/^.+,/, ''));

                                // Close the dialog
                                LoadingDialog.hide();
                            }
                            reader.readAsDataURL(file.src);
                        } else {
                            // Display an error message
                            ctrl.updateValidation(ctrl.el, {
                                isValid: false,
                                invalidMessage: "The file must be a valid image file. Valid types: png, jpg, jpeg, gif"
                            });
                        }
                    });
                });
            }
        }

        // Return the properties
        return props;
    }

    // Displays the edit form
    static edit(itemId: number, onUpdate: () => void) {
        DataSource.List.editForm({
            itemId,
            onCreateEditForm: this.configureForm,
            onUpdate: (item: IAppStoreItem) => {
                // Refresh the item
                DataSource.List.refreshItem(item.Id).then(() => {
                    // Call the update event
                    onUpdate();
                });
            }
        });
    }

    // Displays the new form
    static new(onUpdate: () => void) {
        DataSource.List.newForm({
            onCreateEditForm: this.configureForm,
            onUpdate: (item: IAppStoreItem) => {
                // Refresh the data
                DataSource.List.refreshItem(item.Id).then(() => {
                    // Call the update event
                    onUpdate();
                });
            }
        });
    }

    // Displays the view form
    static view(itemId: number) {
        DataSource.List.viewForm({ itemId });
    }
}