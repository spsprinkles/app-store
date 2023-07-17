import { LoadingDialog } from "dattatable";
import { Components, Helper } from "gd-sprest-bs";
import * as Common from "./common";
import { DataSource, IAppStoreItem } from "./ds";

// Acceptable image file types
const ImageExtensions = [
    ".apng", ".avif", ".bmp", ".cur", ".gif", ".jpg", ".jpeg", ".jfif",
    ".ico", ".pjpeg", ".pjp", ".png", ".svg", ".tif", ".tiff", ".webp"
];

/**
 * Forms
 */
export class Forms {
    // Configures the form
    private static configureForm(props: Components.IListFormEditProps): Components.IListFormEditProps {
        // Set the control rendered event
        props.onControlRendered = (ctrl, fld) => {
            // See if this is a url field
            if (fld.InternalName == "Icon" || fld.InternalName.indexOf("ScreenShot") == 0) {
                // Set a tooltip
                Components.Tooltip({
                    content: "Click to upload an image file.",
                    target: ctrl.textbox.elTextbox
                });

                // Make this textbox read-only
                ctrl.textbox.elTextbox.readOnly = true;

                // Set a click event
                ctrl.textbox.elTextbox.addEventListener("click", () => {
                    // Display a file upload dialog
                    Helper.ListForm.showFileDialog(["image/*"]).then(file => {
                        // Clear the value
                        ctrl.textbox.setValue("");

                        // Get the file name
                        let fileName = file.name.toLowerCase();

                        // Validate the file type
                        if (this.isImageFile(fileName)) {
                            // Show a loading dialog
                            LoadingDialog.setHeader("Reading the File");
                            LoadingDialog.setBody("This will close after the file is converted...");
                            LoadingDialog.show();

                            // Convert the file
                            let reader = new FileReader();
                            reader.onloadend = () => {
                                // Set the value
                                ctrl.textbox.setValue(reader.result as string);

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

    // Displays the detail form
    static detail(item: IAppStoreItem) {
        DataSource.List.viewForm({
            itemId: item.Id,
            onCreateViewForm: props => {
                props.onFormRendered = (form) => {
                    // Get the parent modal dialog element
                    let modalDlg = form.el.closest("div.modal-dialog") as HTMLDivElement;
                    if (modalDlg) {
                        // Change modal to xtra large
                        modalDlg.classList.remove("modal-lg");
                        modalDlg.classList.add("modal-xl");
                        // Hide the footer
                        let footer = modalDlg.querySelector("div.modal-footer") as HTMLDivElement;
                        footer ? footer.classList.add("d-none") : null;
                    }

                    // Clear the form
                    while (form.el.firstChild) { form.el.removeChild(form.el.firstChild); }

                    // Create a new root element
                    let rootEl = document.createElement("div");
                    rootEl.classList.add("container");
                    rootEl.innerHTML = `
                        <div class="row align-items-start">
                            <div class="col-4 mt-3">
                                <div class="row mb-3">
                                    <div class="col d-flex icon justify-content-center"></div>
                                </div>
                                <div class="row">
                                    <div class="col fs-6"><label>App Name:</label> ${item.Title}</div>
                                </div>
                                <div class="row">
                                    <div class="col fs-6"><label>Project Type:</label> ${item.TypeOfProject}</div>
                                </div>
                                <div class="row">
                                    <div class="col fs-6"><label>Description:</label> ${item.Description}</div>
                                </div>
                                <div class="row">
                                    <div class="col fs-6"><label>Additional Information:</label></div>
                                </div>
                            </div>
                            <div class="col-8 screenshots"></div>
                        </div>`;
                    // Add the root element to the form
                    form.el.appendChild(rootEl);

                    // Define the app icon
                    let icon;
                    if (item.Icon) {
                        // Display the image
                        icon = document.createElement("img");
                        icon.style.height = "150px";
                        icon.style.width = "150px";
                        icon.src = item.Icon;
                    } else {
                        // Get the icon by type
                        icon = Common.getIcon(150, 150, item.TypeOfProject);
                    }
                    // Add the app icon to the root element
                    rootEl.querySelector(".icon") ? rootEl.querySelector(".icon").appendChild(icon) : null;

                    // Ensure an Additional Information value exists
                    let addlInfo = rootEl.querySelector("div.col-4 div:last-child div.fs-6");
                    if (item.AdditionalInformation && addlInfo) {
                        // Render the link
                        let elLink = document.createElement("a");
                        elLink.text = (item.AdditionalInformation ? item.AdditionalInformation.Description : "") || "Link";
                        elLink.href = (item.AdditionalInformation ? item.AdditionalInformation.Url : "") || "#";
                        elLink.target = "_blank";
                        elLink.classList.add("ms-1");
                        addlInfo.appendChild(elLink);
                    }

                    // Get each ScreenShot item for the carousel
                    let items: Components.ICarouselItem[] = [];
                    for (let i = 1; i <= 5; i++) {
                        let screenShot = item["ScreenShot" + i] as string;
                        let screenShotAlt = "Screen Shot " + i;
                        if (screenShot) {
                            items.push({ captions: "<h6>" + screenShotAlt + "</h6>", imageAlt: screenShotAlt, imageUrl: screenShot });
                        }
                    }

                    let scrEl = rootEl.querySelector(".screenshots") || rootEl;
                    // Create an image carousel
                    Components.Carousel({
                        el: scrEl,
                        enableControls: true,
                        enableIndicators: true,
                        id: "screenshots" + item.Id,
                        items
                    });
                }

                // Return the properties
                return props;
            }
        });
    }

    // Displays the edit form
    static edit(itemId: number, onUpdate: () => void) {
        DataSource.List.editForm({
            itemId,
            onCreateEditForm: props => { return this.configureForm(props); },
            onUpdate: (item: IAppStoreItem) => {
                // Refresh the item
                DataSource.List.refreshItem(item.Id).then(() => {
                    // Call the update event
                    onUpdate();
                });
            }
        });
    }

    // Determines if the image extension is valid
    private static isImageFile(fileName: string): boolean {
        let isValid = false;

        // Parse the valid file extensions
        for (let i = 0; i < ImageExtensions.length; i++) {
            // See if this is a valid file extension
            if (fileName.endsWith(ImageExtensions[i])) {
                // Set the flag and break from the loop
                isValid = true;
                break;
            }
        }

        // Return the flag
        return isValid;
    }

    // Displays the new form
    static new(onUpdate: () => void) {
        DataSource.List.newForm({
            onCreateEditForm: props => { return this.configureForm(props); },
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
    static view(item: IAppStoreItem) {
        DataSource.List.viewForm({
            itemId: item.Id,
            onCreateViewForm: props => {
                // Customize the screenshots to display the image
                props.onControlRendered = (ctrl, fld) => {
                    // See if this is a url field
                    if (fld.InternalName == "Icon" || fld.InternalName.indexOf("ScreenShot") == 0) {
                        // Clear the element
                        while (ctrl.el.firstChild) { ctrl.el.removeChild(ctrl.el.firstChild); }

                        // Ensure a value exists
                        if (item[fld.InternalName]) {
                            // Display the image
                            let elImage = document.createElement("img");
                            elImage.src = item[fld.InternalName];
                            elImage.style.maxHeight = "250px";
                            ctrl.el.appendChild(elImage);
                        }
                    }
                }

                // Return the properties
                return props;
            }
        });
    }
}