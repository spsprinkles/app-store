import { Documents, LoadingDialog, Modal, DataTable } from "dattatable";
import { Components, Helper, ThemeManager, Types, Web } from "gd-sprest-bs";
import * as jQuery from "jquery";
import * as JSZip from "jszip";
import * as moment from "moment";
import * as Common from "./common";
import { ReadAppLists } from "./readAppLists";
import { DataSource, IAppStoreItem } from "./ds";
import { Security } from "./security";
import Strings from "./strings";

// Acceptable image file types
const ImageExtensions = [
    ".apng", ".avif", ".bmp", ".gif", ".jpg", ".jpeg", ".jfif", ".ico",
    ".pjpeg", ".pjp", ".png", ".svg", ".svgz", ".tif", ".tiff", ".webp", ".xbm"
];

// Flow Action
interface IFlowAction {
    action: string;
    name: string
    type: string;
    value: string;
}

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

        // Add spacing between the tab control and the form
        props.onFormRendered = (form) => {
            form.el ? form.el.classList.add("mt-3") : null;
        }

        // Return the properties
        return props;
    }

    // Displays the form to customize a flow
    static customizeFlowPackage(item: IAppStoreItem, fileInfo: Types.SP.Attachment, onUpdate: () => void) {
        // Clear the modal
        Modal.clear();

        // Set the header
        Modal.setHeader("Customize Flow Package");

        // Show a loading dialog
        LoadingDialog.setHeader("Reading the Package");
        LoadingDialog.setBody("This will close after the flow package is read...");
        LoadingDialog.show();

        // Set the existing values
        let selectedVars = [];
        if (item.FlowData) {
            let items: Components.IDropdownItem[] = JSON.parse(item.FlowData);
            for (let i = 0; i < items.length; i++) {
                let item = items[i];

                // Add the selected variable
                selectedVars.push(item.text);
            }
        }

        // Read the package
        Web(Strings.SourceUrl).getFileByServerRelativeUrl(fileInfo.ServerRelativeUrl).content().execute(data => {
            JSZip.loadAsync(data).then(zipContents => {
                // Parse the files
                zipContents.forEach((path, fileInfo) => {
                    // See if this is the definitions file
                    if (fileInfo.name.endsWith("/definition.json")) {
                        // Get the variables for this definition
                        this.readDefinitionFile(fileInfo).then(variables => {
                            let items: Components.IDropdownItem[] = [];

                            // Parse the variables
                            for (let i = 0; i < variables.length; i++) {
                                let variable = variables[i];

                                // Add the variable
                                items.push({
                                    data: variable,
                                    text: variable.name,
                                    value: variable.name
                                });
                            }

                            // Generate the form
                            let form = Components.Form({
                                el: Modal.BodyElement,
                                controls: [{
                                    label: "Customize Variables",
                                    name: "CustomVariables",
                                    placeholder: "Select Variables...",
                                    description: "Select the variables to customize for the package",
                                    type: Components.FormControlTypes.MultiDropdownCheckbox,
                                    required: true,
                                    items,
                                    value: selectedVars
                                } as Components.IFormControlPropsMultiDropdownCheckbox]
                            });

                            // Generate the footer
                            Components.TooltipGroup({
                                el: Modal.FooterElement,
                                tooltips: [
                                    {
                                        content: "Saves the variables to customize for this package",
                                        btnProps: {
                                            text: "Save",
                                            type: Components.ButtonTypes.OutlinePrimary,
                                            onClick: () => {
                                                // Ensure the form is valid
                                                if (form.isValid()) {
                                                    // Get the selected variables
                                                    let selectedVariables = form.getValues()["CustomVariables"];

                                                    // Show a loading dialog
                                                    LoadingDialog.setHeader("Updating the Item");
                                                    LoadingDialog.setBody("This will close after the item is updated...");
                                                    LoadingDialog.show();

                                                    // Update the item
                                                    item.update({
                                                        FlowData: JSON.stringify(selectedVariables)
                                                    }).execute(
                                                        () => {
                                                            // Call the update event
                                                            onUpdate();

                                                            // Close the form
                                                            Modal.hide();

                                                            // Hide the loading dialog
                                                            LoadingDialog.hide();
                                                        }
                                                    )
                                                }
                                            }
                                        }
                                    },
                                    {
                                        content: "Closes the modal",
                                        btnProps: {
                                            text: "Close",
                                            type: Components.ButtonTypes.OutlinePrimary,
                                            onClick: () => {
                                                // Hide the form
                                                Modal.hide();
                                            }
                                        }
                                    }
                                ]
                            })

                            // Hide the loading dialog
                            LoadingDialog.hide();

                            // Show the modal
                            Modal.show();
                        });
                    }
                });
            });
        }, () => {
            // Error loading the package
            Modal.setBody("Error loading the power automate package...");

            // Hide the loading dialog
            LoadingDialog.hide();

            // Show the modal
            Modal.show();
        });
    }

    // Displays the edit form
    static edit(item: IAppStoreItem, onUpdate: () => void) {
        DataSource.List.editForm({
            itemId: item.Id,
            onCreateEditForm: props => { return this.configureForm(props); },
            onUpdate: (updatedItem: IAppStoreItem) => {
                // Refresh the item
                DataSource.List.refreshItem(updatedItem.Id).then(refreshedItem => {
                    // See if the item is approved
                    if (item.Status != "Approved" && refreshedItem.Status == "Approved") {
                        // Add the developers and then call the update event
                        this.addDevelopers(refreshedItem).then(onUpdate);
                    } else {
                        // Call the update event
                        onUpdate();
                    }
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
                        excludeFields: [
                            "Attachments", "ScreenShot1", "ScreenShot2", "ScreenShot3",
                            "ScreenShot4", "ScreenShot5", "VideoURL"
                        ]
                    },
                    {
                        title: "Screen Shots",
                        fields: [
                            "ScreenShot1", "ScreenShot2", "ScreenShot3",
                            "ScreenShot4", "ScreenShot5", "VideoURL"
                        ],
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
                        fields: [],
                        onRendered: (el) => {
                            // Render the form
                            ReadAppLists.renderForm(el, item);

                            // Render the footer
                            ReadAppLists.renderFooter(el, item);
                        }
                    }
                ]
            }
        });
    }

    // Displays the form to generate a custom flow package
    static generateFlow(item: IAppStoreItem, fileInfo: Types.SP.Attachment) {
        // Clear the modal
        Modal.clear();

        // Set the header
        Modal.setHeader("Generate Flow Package");

        // Parse the variables
        let controls: Components.IFormControlProps[] = [];
        let items: Components.IDropdownItem[] = JSON.parse(item.FlowData);
        let variables: IFlowAction[] = [];
        for (let i = 0; i < items.length; i++) {
            let item = items[i];
            let variable: IFlowAction = item.data;
            let isTextbox = variable.type == "string" || variable.type.startsWith("int");

            // Save a reference to the variable
            variables.push(variable);

            // Generate a form control
            controls.push({
                data: variable,
                name: variable.name,
                label: variable.name,
                type: isTextbox ? Components.FormControlTypes.TextField : Components.FormControlTypes.TextArea,
                value: variable.value
            });
        }

        // Generate the form
        let form = Components.Form({
            el: Modal.BodyElement,
            controls
        });

        // Set the footer
        Components.TooltipGroup({
            el: Modal.FooterElement,
            tooltips: [
                {
                    content: "Generates the flow package.",
                    btnProps: {
                        text: "Generate",
                        type: Components.ButtonTypes.OutlinePrimary,
                        onClick: () => {
                            // Ensure the form is valid
                            if (form.isValid()) {
                                let values = form.getValues();

                                // Show a loading dialog
                                LoadingDialog.setHeader("Generating Package");
                                LoadingDialog.setBody("This dialog will close after the package is generated...");
                                LoadingDialog.show();

                                // Parse the variables
                                for (let i = 0; i < variables.length; i++) {
                                    let variable = variables[i];

                                    // Set the value
                                    variable.value = values[variable.name];
                                }

                                // Generate the package
                                this.generateFlowPackage(fileInfo, variables).then(zipFile => {
                                    // Generate the new file
                                    zipFile.generateAsync({ type: "uint8array" }).then(contents => {
                                        // See if this is IE or Mozilla
                                        if (Blob && navigator && navigator["msSaveBlob"]) {
                                            // Download the file
                                            navigator["msSaveBlob"](new Blob([contents], { type: "octet/stream" }), fileInfo.FileName);
                                        } else {
                                            // Generate an anchor
                                            var anchor = document.createElement("a");
                                            anchor.download = fileInfo.FileName;
                                            anchor.href = URL.createObjectURL(new Blob([contents], { type: "octet/stream" }));
                                            anchor.target = "__blank";

                                            // Download the file
                                            anchor.click();
                                        }

                                        // Close the loading dialog
                                        LoadingDialog.hide();
                                    }, () => {
                                        // Error loading the package
                                        Modal.setBody("Error generating the power automate package...");

                                        // Show the modal
                                        Modal.show();
                                    });
                                });
                            }
                        }
                    }
                },
                {
                    content: "Closes the form.",
                    btnProps: {
                        text: "Close",
                        type: Components.ButtonTypes.OutlinePrimary,
                        onClick: () => {
                            // Close the modal
                            Modal.hide();
                        }
                    }
                }
            ]
        });

        // Show the modal
        Modal.show();
    }

    // Generates the flow package
    private static generateFlowPackage(fileInfo: Types.SP.Attachment, variables: IFlowAction[]): PromiseLike<JSZip> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Read the package
            Web(Strings.SourceUrl).getFileByServerRelativeUrl(fileInfo.ServerRelativeUrl).content().execute(data => {
                JSZip.loadAsync(data).then(zipContents => {
                    // Parse the files
                    zipContents.forEach((path, fileInfo) => {
                        // See if this is the definitions file
                        if (fileInfo.name.endsWith("/definition.json")) {
                            // Read the file contents
                            fileInfo.async("string").then(contents => {
                                // Get the actions
                                let def = JSON.parse(contents);
                                let actions = def?.properties?.definition?.actions || {};

                                // Parse the variables
                                for (let i = 0; i < variables.length; i++) {
                                    let variable = variables[i];

                                    // Ensure the variable exists for this action
                                    if (actions[variable.action] && actions[variable.action].inputs?.variables) {
                                        // Update the value
                                        actions[variable.action].inputs.variables[0].value = variable.value;
                                    }
                                }

                                // Update the file
                                zipContents.file(fileInfo.name, JSON.stringify(def));

                                // Resolve the request
                                resolve(zipContents);
                            });
                        }
                    });
                }, reject);
            }, reject);
        });
    }

    // Determines if package is a flow
    static isFlowPackage(attachment: Types.SP.Attachment): PromiseLike<{ attachment: Types.SP.Attachment; isFlow: boolean }> {
        let filePath = attachment.ServerRelativeUrl;

        // Return a promise
        return new Promise(resolve => {
            // See if this is not a zip file
            if (!filePath.endsWith(".zip")) { resolve({ attachment, isFlow: false }); return; }

            // Read the package
            Web(Strings.SourceUrl).getFileByServerRelativeUrl(filePath).content().execute(data => {
                JSZip.loadAsync(data).then(zipContents => {
                    // See if this has a manifest file in the root
                    resolve({ attachment, isFlow: zipContents.files["manifest.json"] ? true : false });
                });
            });
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
                        excludeFields: [
                            "Attachments", "ScreenShot1", "ScreenShot2", "ScreenShot3",
                            "ScreenShot4", "ScreenShot5", "VideoURL"
                        ]
                    },
                    {
                        title: "Screen Shots",
                        fields: [
                            "ScreenShot1", "ScreenShot2", "ScreenShot3",
                            "ScreenShot4", "ScreenShot5", "VideoURL"
                        ],
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
            onSetHeader: (el) => {
                el.querySelector("h5") ? el.querySelector("h5").innerHTML = "New App Request" : null;
            },
            onUpdate: (item: IAppStoreItem) => {
                // Refresh the data
                DataSource.RequestsList.refreshItem(item.Id).then(onUpdate);
            }
        });
    }

    // Reads the flow definition for variables
    private static readDefinitionFile(fileInfo: JSZip.JSZipObject): PromiseLike<IFlowAction[]> {
        let variables: IFlowAction[] = [];

        // Return a promise
        return new Promise(resolve => {
            // Read the file contents
            fileInfo.async("string").then(contents => {
                // Convert to an object
                let def = JSON.parse(contents);

                // Parse the actions
                let actions = def?.properties?.definition?.actions || {};
                for (let key in actions) {
                    let action = actions[key];

                    // Get the variable name
                    let variable = (action?.inputs?.variables || [])[0];
                    if (variable?.name) {
                        variables.push({
                            action: key,
                            name: variable.name,
                            type: variable.type,
                            value: variable.value
                        });
                    }
                }

                // Resolve the request
                resolve(variables);
            });
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
            ThemeManager.IsInverted ? icon.classList.add("invert") : null;
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
                props.isDark = ThemeManager.IsInverted;
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
        Modal.setHeader("App Requests");

        // Render the table
        new DataTable({
            el: Modal.BodyElement,
            rows: DataSource.RequestsList.Items,
            dtProps: {
                dom: 'rt<"row"<"col-sm-4"l><"col-sm-4"i><"col-sm-4"p>>',
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
                language: {
                    emptyTable: "No app requests exist",
                },
                // Sort descending by Start Date
                order: [[1, "asc"]],
                pageLength: 10
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
            content: "Close the dialog",
            btnProps: {
                text: "Close",
                type: Components.ButtonTypes.OutlineSecondary,
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