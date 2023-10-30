import { Documents, LoadingDialog, Modal } from "dattatable";
import { Components, Helper, Web } from "gd-sprest-bs";
import * as moment from "moment";
import * as Common from "./common";
import { DataSource, IAppStoreItem } from "./ds";
import { Security } from "./security";
import Strings from "./strings";

// Acceptable image file types
const ImageExtensions = [
    ".apng", ".avif", ".bmp", ".gif", ".jpg", ".jpeg", ".jfif", ".ico",
    ".pjpeg", ".pjp", ".png", ".svg", ".svgz", ".tif", ".tiff", ".webp", ".xbm"
];

/**
 * Forms
 */
export class Forms {
    // Adds the developers to the group
    private static addDevelopers(item: IAppStoreItem) {
        // Return a promise
        return new Promise((resolve, reject) => {
            let group = Web().SiteGroups().getById(Security.DeveloperGroup.Id);

            // Parse the owners and add them to the developers group
            Helper.Executor(item.Developers.results, user => {
                // See if the user is not in the group
                if (!Security.isDeveloper(user.Id)) {
                    // Add them to the group
                    group.Users().addUserById(user.Id).batch();
                }
            }).then(() => {
                // Execute the batch request
                group.execute(resolve, reject);
            });
        });
    }

    // Configures the form
    private static configureForm(props: Components.IListFormEditProps): Components.IListFormEditProps {
        // Include the attachments
        props.displayAttachments = true;

        // Set the control rendered event
        props.onControlRendered = (ctrl, fld) => {
            // See if this is a url field
            if (fld.InternalName && (fld.InternalName == "Icon" || fld.InternalName.indexOf("ScreenShot") == 0)) {
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
                                invalidMessage: "The file must be a valid image file. Valid types: png, svg, jpg, gif"
                            });
                        }
                    });
                });
            }
        }

        props.onFormRendered = (form) => {
            let col12;
            let body = form.el.closest(".modal-body");
            col12 = body.querySelector(".row>.col-12");
            (body && col12) ? col12.classList.add("mb-3") : null;
        }

        // Return the properties
        return props;
    }

    // Displays the edit form
    static edit(itemId: number, onUpdate: () => void) {
        DataSource.List.editForm({
            itemId,
            onCreateEditForm: props => { return this.configureForm(props); },
            onUpdate: (item: IAppStoreItem) => {
                // Refresh the item
                DataSource.List.refreshItem(item.Id).then(updatedItem => {
                    // Add the developers
                    this.addDevelopers(updatedItem).then(onUpdate);
                });
            },
            tabInfo: {
                tabs: [
                    {
                        title: "App Details",
                        excludeFields: ["Attachments"]
                    },
                    {
                        title: "Attachments",
                        fields: [],
                        onRendered: (el) => {
                            // Render the attachments
                            new Documents({
                                el,
                                listName: Strings.Lists.Main,
                                itemId: itemId
                            });
                        }
                    }
                ]
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
                DataSource.List.refreshItem(item.Id).then(newItem => {
                    // Add the developers
                    this.addDevelopers(newItem).then(onUpdate);
                });
            },
            tabInfo: {
                tabs: [
                    {
                        title: "App Details",
                        excludeFields: ["Attachments"]
                    },
                    {
                        title: "Attachments",
                        fields: ["Attachments"]
                    }
                ]
            }
        });
    }

    // Displays the view form
    static view(item: IAppStoreItem) {
        // Clear the modal
        Modal.clear();

        // Set the modal props
        Modal.setType(Components.ModalTypes.XLarge);

        // Set the header
        Modal.setHeader("");

        // Hide the footer
        Modal.FooterElement.classList.add("d-none");

        // Ensure developers exist
        let developers = "";
        if (item.Developers && item.Developers.results) {
            // Parse the devs
            let devArr = [];
            item.Developers.results.forEach(dev => {
                // Append the title
                devArr.push(dev.Title);
            });

            // Display the devs
            developers = devArr.join(", ");
        }

        // See if this is from the app catalog
        let moreInfo = "";
        if (item.IsAppCatalogItem) {
            // Set the more info link
            moreInfo = `<a href="${DataSource.AppCatalogUrl}?app-id=${item.Id}" target="_blank">App Dashboard (Details)</a>`;
        } else {
            // Render the link
            item.MoreInfo ? moreInfo = `<a href="${item.MoreInfo.Url}" target="_blank">${item.MoreInfo.Description}</a>` : null;
        }

        // Ensure a SupportURL value exists
        let support = "";
        item.SupportURL ? support = `<a href="${item.SupportURL.Url}" target="_blank">${item.SupportURL.Description}</a>` : null;

        // Generate the attachments
        let attachments = "";
        if (item.AttachmentFiles && item.AttachmentFiles.results) {
            // Parse the attachments
            for (let i = 0; i < item.AttachmentFiles.results.length; i++) {
                let attachment = item.AttachmentFiles.results[i];

                // Add the link
                attachments += Common.isWopi(attachment.FileName) ? `<br/><a href="${Strings.SourceUrl}/_layouts/15/WopiFrame.aspx?sourcedoc=${attachment.ServerRelativeUrl}&action=view" target="_blank">${attachment.FileName}</a>` : `<br/><a href="${attachment.ServerRelativeUrl}" target="_blank">${attachment.FileName}</a>`;
            }
        }

        // Create a new div element
        let div = document.createElement("div");
        div.classList.add("container-fluid");
        div.innerHTML = `
            <div class="row align-items-start">
                <div class="col-4 mt-3">
                    <div class="row mb-3">
                        <div class="col d-flex icon justify-content-center"></div>
                    </div>
                    <div class="row">
                        <div class="col fs-6"><label>App Name:</label>&nbsp;${item.Title}</div>
                    </div>
                    <div class="row">
                        <div class="col fs-6"><label>App Type:</label>&nbsp;${item.AppType}</div>
                    </div>
                    <div class="row">
                        <div class="col fs-6"><label>Description:</label>&nbsp;${item.Description}</div>
                    </div>
                    <div class="row">
                        <div class="col fs-6"><label>Developers:</label>&nbsp;${developers}</div>
                    </div>
                    <div class="row">
                        <div class="col fs-6"><label>Organization:</label>&nbsp;${item.Organization}</div>
                    </div>
                    <div class="row">
                        <div class="col fs-6"><label>More Info:</label>&nbsp;${moreInfo}</div>
                    </div>
                    <div class="row">
                        <div class="col fs-6"><label>Support:</label>&nbsp;${support}</div>
                    </div>
                    <div class="row">
                        <div class="col fs-6"><label>Updated:</label>&nbsp;${moment(item.Modified).format(Strings.DateFormat)}</div>
                    </div>
                    <div class="row">
                        <div class="col fs-6"><label>Attachments:</label>${attachments}</div>
                    </div>
                </div>
                <div class="col-8 screenshots"></div>
            </div>`;

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
            icon = Common.getIcon(150, 150, item.AppType);
        }
        // Add the app icon to the div element
        div.querySelector(".icon") ? div.querySelector(".icon").appendChild(icon) : null;

        // Get each ScreenShot item for the carousel
        let items: Components.ICarouselItem[] = [];
        for (let i = 1; i <= 5; i++) {
            let screenShot = item["ScreenShot" + i] as string;
            let screenShotAlt = "Screen Shot " + i;
            if (screenShot) {
                items.push({ captions: "<h6>" + screenShotAlt + "</h6>", imageAlt: screenShotAlt, imageUrl: screenShot });
            }
        }

        // Create an image carousel
        let scrEl = (div.querySelector(".screenshots") || div) as HTMLElement;
        Components.Carousel({
            el: scrEl,
            enableControls: true,
            enableIndicators: true,
            id: "screenshots" + item.Id,
            items
        });

        // Add the div element to the form
        Modal.BodyElement.appendChild(div);

        // Show the modal
        Modal.show();
    }
}