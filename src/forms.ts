import { Documents, LoadingDialog, Modal, DataTable } from "dattatable";
import { Components, Helper } from "gd-sprest-bs";
import * as jQuery from "jquery";
import * as moment from "moment";
import * as Common from "./common";
import { CreateTemplate } from "./createTemplate";
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
            let requests = [];

            // Parse the owners and add them to the developers group
            Helper.Executor(item.Developers.results, user => {
                // Add the developer
                requests.push(Security.addDeveloper(user.Id));
            });

            // Wait for the requests to complete
            Promise.all(requests).then(resolve, reject);
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
                    options: {
                        theme: "sharepoint"
                    },
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
            let closeBtn;
            let col12;
            let modal = form.el.closest(".modal-content");
            modal ? col12 = modal.querySelector(".modal-body .row>.col-12") : null;
            modal ? closeBtn = modal.querySelector(".modal-header .btn-close") : null;
            (modal && col12) ? col12.classList.add("mb-3") : null;
            (closeBtn && DataSource.ThemeInfo && DataSource.ThemeInfo.isInverted) ? closeBtn.classList.add("invert") : null;
        }

        // Return the properties
        return props;
    }

    // Displays the edit form
    static edit(item: IAppStoreItem, onUpdate: () => void) {
        DataSource.List.editForm({
            itemId: item.Id,
            onCreateEditForm: props => { return this.configureForm(props); },
            onUpdate: (item: IAppStoreItem) => {
                // Refresh the item
                DataSource.List.refreshItem(item.Id).then(updatedItem => {
                    // Add the developers
                    this.addDevelopers(updatedItem).then(onUpdate);
                });
            },
            tabInfo: {
                onClick: (el, tab) => {
                    // See if the App Details tab was not clicked
                    if (tab.tabName != "App Details") {
                        Modal.FooterElement.classList.add('d-none');
                    } else {
                        Modal.FooterElement.classList.remove('d-none');
                    }
                },
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
                                itemId: item.Id
                            });
                        }
                    },
                    {
                        title: "Template",
                        onRendered: (el) => {
                            // Render the form
                            CreateTemplate.renderForm(el);

                            // Render the footer
                            CreateTemplate.renderFooter(el, item, false);
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

    // Displays the new request form
    static newRequest(onUpdate: () => void) {
        DataSource.RequestsList.newForm({
            onUpdate: (item: IAppStoreItem) => {
                // Refresh the data
                DataSource.RequestsList.refreshItem(item.Id).then(onUpdate);
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

        let closeBtn;
        Modal.HeaderElement ? closeBtn = Modal.HeaderElement.closest(".modal-header").querySelector(".btn-close") : null;
        (closeBtn && DataSource.ThemeInfo && DataSource.ThemeInfo.isInverted) ? closeBtn.classList.add("invert") : null;

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
                        <div class="attachments col fs-6"></div>
                    </div>
                </div>
                <div class="col-8 screenshots"></div>
            </div>`;

        // Define the app icon
        let icon;
        if (item.Icon) {
            // Display the image
            icon = document.createElement("img");
            (DataSource.ThemeInfo && DataSource.ThemeInfo.isInverted) ? icon.classList.add("invert") : null;
            icon.style.height = "150px";
            icon.style.width = "150px";
            icon.src = item.Icon;
        } else {
            // Get the icon by type
            icon = Common.getIcon(150, 150, item.AppType, 'icon-type');
        }
        // Add the app icon to the div element
        div.querySelector(".icon") ? div.querySelector(".icon").appendChild(icon) : null;

        // Generate the attachments
        if (item.AttachmentFiles && item.AttachmentFiles.results && item.AttachmentFiles.results.length > 0) {
            let attachments = "<label>Attachments:</label>";
            // Parse the attachments
            for (let i = 0; i < item.AttachmentFiles.results.length; i++) {
                let attachment = item.AttachmentFiles.results[i];

                // Add the link
                attachments += Common.isWopi(attachment.FileName) ? `<br/><a href="${Strings.SourceUrl}/_layouts/15/WopiFrame.aspx?sourcedoc=${attachment.ServerRelativeUrl}&action=view" target="_blank">${attachment.FileName}</a>` : `<br/><a href="${attachment.ServerRelativeUrl}" target="_blank">${attachment.FileName}</a>`;
            }
            // Add the app icon to the div element
            div.querySelector(".attachments") ? div.querySelector(".attachments").innerHTML = attachments : null;
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
        let scrEl = (div.querySelector(".screenshots") || div) as HTMLElement;
        Components.Carousel({
            el: scrEl,
            enableControls: true,
            enableIndicators: true,
            id: "screenshots" + item.Id,
            items,
            onRendered: (el, props) => {
                DataSource.ThemeInfo ? props.isDark = DataSource.ThemeInfo.isInverted : null;
            }
        });

        // Add the div element to the form
        Modal.BodyElement.appendChild(div);

        // Show the modal
        Modal.show();
    }

    // Displays the view the current requests
    static viewRequests() {
        // Clear the modal
        Modal.clear();

        // Set the modal props
        Modal.setType(Components.ModalTypes.XLarge);

        // Set the header
        Modal.setHeader("View Requests");

        // Update the close button
        let closeBtn;
        Modal.HeaderElement ? closeBtn = Modal.HeaderElement.closest(".modal-header").querySelector(".btn-close") : null;
        (closeBtn && DataSource.ThemeInfo && DataSource.ThemeInfo.isInverted) ? closeBtn.classList.add("invert") : null;

        // Render the table
        new DataTable({
            el: Modal.BodyElement,
            rows: DataSource.RequestsList.Items,
            dtProps: {
                dom: 'rt<"row"<"col-sm-4"l><"col-sm-4"i><"col-sm-4"p>>',
                pageLength: 10,
                createdRow: function (row, data, index) {
                    jQuery('td', row).addClass('align-middle');
                },
                // Add some classes to the dataTable elements
                drawCallback: function (settings) {
                    let api = new jQuery.fn.dataTable.Api(settings) as any;
                    let div = api.table().container() as HTMLDivElement;
                    let table = api.table().node() as HTMLTableElement;
                    div.querySelector(".dataTables_info").classList.add("text-center");
                    div.querySelector(".dataTables_length").classList.add("pt-2");
                    div.querySelector(".dataTables_paginate").classList.add("pt-03");
                    table.classList.remove("no-footer");
                    table.classList.add("tbl-footer");
                    table.classList.add("table-striped");
                },
                headerCallback: function (thead, data, start, end, display) {
                    jQuery('th', thead).addClass('align-middle');
                },
                // Sort descending by Start Date
                order: [[1, "asc"]],
                language: {
                    emptyTable: "No app requests exist",
                },
            },
            columns: [
                {
                    name: "Title",
                    title: "App Name"
                },
                {
                    name: "AppType",
                    title: "App Type"
                },
                {
                    name: "Status",
                    title: "App Status"
                },
                {
                    name: "Description",
                    title: "Description"
                },
                {
                    name: "Organization",
                    title: "Organization"
                },
                {
                    name: "Developers",
                    title: "Developers",
                    onRenderCell: (el, column, item: IAppStoreItem) => {
                        let devs = item.Developers && item.Developers.results || [];

                        // Parse the devs
                        let strDevs = [];
                        devs.forEach(dev => {
                            // Append the title
                            strDevs.push(dev.Title);
                        });

                        // Display the devs
                        el.innerHTML = `${strDevs.join("<br/>")}`;
                        el.setAttribute("data-filter", strDevs.join(","));
                    }
                }
            ]
        });

        // Render the footer
        Components.Tooltip({
            el: Modal.FooterElement,
            content: "Closes the dialog",
            btnProps: {
                text: "Close",
                type: Components.ButtonTypes.OutlinePrimary,
                onClick: () => {
                    // Hide the modal
                    Modal.hide();
                }
            }
        });

        // Show the modal
        Modal.show();
    }
}