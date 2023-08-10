import { Documents, LoadingDialog, Modal } from "dattatable";
import { Components, Helper } from "gd-sprest-bs";
import * as moment from "moment";
import * as Common from "./common";
import { DataSource, IAppStoreItem } from "./ds";
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
                DataSource.List.refreshItem(item.Id).then(() => {
                    // Call the update event
                    onUpdate();
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
                DataSource.List.refreshItem(item.Id).then(() => {
                    // Call the update event
                    onUpdate();
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

        // Generate the attachments
        let attachments = "";
        if (item.AttachmentFiles && item.AttachmentFiles.results) {
            // Parse the attachments
            debugger;
            for (let i = 0; i < item.AttachmentFiles.results.length; i++) {
                let attachment = item.AttachmentFiles.results[i];

                // Add the link
                attachments += `<br/><a href="${attachment.ServerRelativeUrl}">${attachment.FileName}</a>`;
            }
        }

        // Create a new div element
        let div = document.createElement("div");
        div.classList.add("container");
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
                        <div class="col fs-6 developers"><label>Developers:</label>&nbsp;</div>
                    </div>
                    <div class="row">
                        <div class="col fs-6"><label>Organization:</label>&nbsp;${item.Organization}</div>
                    </div>
                    <div class="row">
                        <div class="col fs-6 more-info"><label>More Info:</label></div>
                    </div>
                    <div class="row">
                        <div class="col fs-6 support"><label>Support:</label></div>
                    </div>
                    <div class="row">
                        <div class="col fs-6"><label>Updated:</label>&nbsp;${moment(item.Modified).format(Strings.DateFormat)}</div>
                    </div>
                    <div class="row">
                        <div class="col fs-6 attachments"><label>Attachments:</label>${attachments}</div>
                    </div>
                </div>
                <div class="col-8 screenshots"></div>
            </div>`;

        // Add the div element to the form
        Modal.BodyElement.appendChild(div);

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

        // Ensure developers exist
        let devDiv = div.querySelector("div.developers");
        if (item.Developers && item.Developers.results && devDiv) {
            // Parse the devs
            let strDevs = [];
            item.Developers.results.forEach(dev => {
                // Append the title
                strDevs.push(dev.Title);
            });

            // Display the devs
            devDiv.append(strDevs.join(", "));
        }

        // Ensure a more info value exists
        let moreInfo = div.querySelector("div.more-info");
        if (item.MoreInfo && moreInfo) {
            // Render the link
            let elLink = document.createElement("a");
            elLink.text = (item.MoreInfo ? item.MoreInfo.Description : "") || "Link";
            elLink.href = (item.MoreInfo ? item.MoreInfo.Url : "") || "#";
            elLink.target = "_blank";
            elLink.classList.add("ms-1");
            moreInfo.appendChild(elLink);
        }

        // Ensure a SupportURL value exists
        let support = div.querySelector("div.support");
        if (item.SupportURL && support) {
            // Render the link
            let elLink = document.createElement("a");
            elLink.text = (item.SupportURL ? item.SupportURL.Description : "") || "Link";
            elLink.href = (item.SupportURL ? item.SupportURL.Url : "") || "#";
            elLink.target = "_blank";
            elLink.classList.add("ms-1");
            support.appendChild(elLink);
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

        // Create an image carousel
        let scrEl = div.querySelector(".screenshots") || div;
        Components.Carousel({
            el: scrEl,
            enableControls: true,
            enableIndicators: true,
            id: "screenshots" + item.Id,
            items
        });

        // Show the modal
        Modal.show();
    }
}