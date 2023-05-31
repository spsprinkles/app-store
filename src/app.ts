import { Dashboard, LoadingDialog } from "dattatable";
import { Components, Helper } from "gd-sprest-bs";
import * as jQuery from "jquery";
import { DataSource, IAppStoreItem } from "./ds";
import Strings from "./strings";

/**
 * Main Application
 */
export class App {
    // Constructor
    constructor(el: HTMLElement) {
        // Render the dashboard
        this.render(el);
    }

    // Configures the form
    private configureForm(props: Components.IListFormEditProps): Components.IListFormEditProps {
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

    // Renders the dashboard
    private render(el: HTMLElement) {
        // Create the dashboard
        let dashboard = new Dashboard({
            el,
            hideHeader: true,
            useModal: true,
            filters: {
                items: [{
                    header: "By Project Type",
                    items: DataSource.FiltersTypeOfProject,
                    onFilter: (value: string) => {
                        // Filter the table
                        dashboard.filter(0, value);
                    }
                }]
            },
            navigation: {
                title: Strings.ProjectName,
                items: [
                    {
                        className: "btn-outline-light",
                        text: "Create Item",
                        isButton: true,
                        onClick: () => {
                            // Show the new form
                            DataSource.List.newForm({
                                onCreateEditForm: this.configureForm,
                                onUpdate: (item: IAppStoreItem) => {
                                    // Refresh the data
                                    DataSource.List.refreshItem(item.Id).then(() => {
                                        // Refresh the table
                                        dashboard.refresh(DataSource.List.Items);
                                    });
                                }
                            });
                        }
                    }
                ]
            },
            footer: {
                itemsEnd: [
                    {
                        text: "v" + Strings.Version
                    }
                ]
            },
            table: {
                rows: DataSource.List.Items,
                dtProps: {
                    dom: 'rt<"row"<"col-sm-4"l><"col-sm-4"i><"col-sm-4"p>>',
                    columnDefs: [
                        {
                            "targets": 4,
                            "orderable": false,
                            "searchable": false
                        }
                    ],
                    createdRow: function (row, data, index) {
                        jQuery('td', row).addClass('align-middle');
                    },
                    drawCallback: function (settings) {
                        let api = new jQuery.fn.dataTable.Api(settings) as any;
                        jQuery(api.context[0].nTable).removeClass('no-footer');
                        jQuery(api.context[0].nTable).addClass('tbl-footer');
                        jQuery(api.context[0].nTable).addClass('table-striped');
                        jQuery(api.context[0].nTableWrapper).find('.dataTables_info').addClass('text-center');
                        jQuery(api.context[0].nTableWrapper).find('.dataTables_length').addClass('pt-2');
                        jQuery(api.context[0].nTableWrapper).find('.dataTables_paginate').addClass('pt-03');
                    },
                    headerCallback: function (thead, data, start, end, display) {
                        jQuery('th', thead).addClass('align-middle');
                    },
                    // Order by the 1st column by default; ascending
                    order: [[1, "asc"]]
                },
                columns: [
                    {
                        name: "TypeOfProject",
                        title: "Project Type"
                    },
                    {
                        name: "Title",
                        title: "App Name",
                    },
                    {
                        name: "Description",
                        title: "Description"
                    },
                    {
                        name: "",
                        title: "Additional Information",
                        onRenderCell: (el, column, item: IAppStoreItem) => {
                            // Render the link
                            let elLink = document.createElement("a");
                            elLink.text = "Additional Information";
                            elLink.href = item.AdditionalInformation ? item.AdditionalInformation.Url : "";
                            elLink.target = "_blank";
                            el.appendChild(elLink);
                        }
                    },
                    {
                        name: "",
                        title: "",
                        onRenderCell: (el, column, item: IAppStoreItem) => {
                            // Render the actions
                            Components.TooltipGroup({
                                el,
                                isSmall: true,
                                tooltips: [
                                    {
                                        content: "Click to view the item.",
                                        btnProps: {
                                            text: "View",
                                            type: Components.ButtonTypes.OutlinePrimary,
                                            onClick: () => {
                                                // Edit the item
                                                DataSource.List.viewForm({
                                                    itemId: item.Id
                                                });
                                            }
                                        }
                                    },
                                    {
                                        content: "Click to edit the item.",
                                        btnProps: {
                                            text: "Edit",
                                            type: Components.ButtonTypes.OutlinePrimary,
                                            onClick: () => {
                                                // Edit the item
                                                DataSource.List.editForm({
                                                    itemId: item.Id,
                                                    onCreateEditForm: this.configureForm,
                                                    onUpdate: (item: IAppStoreItem) => {
                                                        // Refresh the item
                                                        DataSource.List.refreshItem(item.Id).then(() => {
                                                            // Refresh the table
                                                            dashboard.refresh(DataSource.List.Items);
                                                        });
                                                    }
                                                });
                                            }
                                        }
                                    }
                                ]
                            });
                        },
                    }
                ]
            }
        });
    }
}