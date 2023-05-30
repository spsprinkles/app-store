import { Dashboard } from "dattatable";
import * as jQuery from "jquery";
import { DataSource, IAppStoreItem } from "./ds";
import Strings from "./strings";
import { Components, SPTypes } from "gd-sprest-bs";

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
            if (fld.FieldTypeKind == SPTypes.FieldType.URL) {
                // Set the option
                (ctrl as Components.IFormControlUrlProps).showDescription = false;
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
                            Components.Button({
                                el,
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
                                    })
                                }
                            })
                        },
                    }
                ]
            }
        });
    }
}