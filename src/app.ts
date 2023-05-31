import { Dashboard } from "dattatable";
import { Components } from "gd-sprest-bs";
import { appIndicator } from "gd-sprest-bs/build/icons/svgs/appIndicator";
import { filterSquare } from "gd-sprest-bs/build/icons/svgs/filterSquare";
import { plusSquare } from "gd-sprest-bs/build/icons/svgs/plusSquare";
import * as jQuery from "jquery";
import { DataSource, IAppStoreItem } from "./ds";
import { Forms } from "./forms";
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
            footer: {
                itemsEnd: [
                    {
                        className: "p-0 pe-none text-dark",
                        text: "v" + Strings.Version
                    }
                ]
            },
            navigation: {
                // Add the branding icon & text
                onRendering: (props) => {
                    // Set the class names
                    props.className = "bg-sharepoint navbar-expand rounded-top";

                    // Set the brand
                    let brand = document.createElement("div");
                    brand.className = "d-flex";
                    brand.appendChild(appIndicator());
                    brand.append(Strings.ProjectName);
                    brand.querySelector("svg").classList.add("me-75");
                    props.brand = brand;
                },
                // Adjust the brand alignment
                onRendered: (el) => {
                    el.querySelector("nav div.container-fluid").classList.add("ps-3");
                    el.querySelector("nav div.container-fluid a.navbar-brand").classList.add("pe-none");
                },
                onSearchRendered: (el) => {
                    el.setAttribute("placeholder", "Find an app");
                },
                showFilter: false
            },
            subNavigation: {
                itemsEnd: [
                    {
                        text: "Add an App",
                        onRender: (el, item) => {
                            // Clear the existing button
                            el.innerHTML = "";
                            // Create a span to wrap the icon in
                            let span = document.createElement("span");
                            span.className = "bg-white d-inline-flex ms-2 rounded";
                            el.appendChild(span);

                            // Render a tooltip
                            Components.Tooltip({
                                el: span,
                                content: item.text,
                                btnProps: {
                                    // Render the icon button
                                    className: "p-1 pe-2",
                                    iconClassName: "me-1",
                                    iconType: plusSquare,
                                    iconSize: 24,
                                    isSmall: true,
                                    text: "Add",
                                    type: Components.ButtonTypes.OutlineSecondary,
                                    onClick: () => {
                                        // Show the new form
                                        Forms.new(() => {
                                            // Refresh the table
                                            dashboard.refresh(DataSource.List.Items);
                                        });
                                    }
                                },
                            });
                        }
                    },
                    {
                        text: "Filters",
                        onRender: (el, item) => {
                            // Clear the existing button
                            el.innerHTML = "";
                            // Create a span to wrap the icon in
                            let span = document.createElement("span");
                            span.className = "bg-white d-inline-flex ms-2 rounded";
                            el.appendChild(span);

                            // Render a tooltip
                            Components.Tooltip({
                                el: span,
                                content: "Show " + item.text,
                                btnProps: {
                                    // Render the icon button
                                    className: "p-1 pe-2",
                                    iconClassName: "me-1",
                                    iconType: filterSquare,
                                    iconSize: 24,
                                    isSmall: true,
                                    text: item.text,
                                    type: Components.ButtonTypes.OutlineSecondary,
                                    onClick: () => {
                                        // Show the filter panel
                                        dashboard.showFilter();
                                    }
                                },
                            });
                        }
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
                            // Ensure a value exists
                            if (item.AdditionalInformation) {
                                // Render the link
                                let elLink = document.createElement("a");
                                elLink.text = "Additional Information";
                                elLink.href = item.AdditionalInformation ? item.AdditionalInformation.Url : "";
                                elLink.target = "_blank";
                                el.appendChild(elLink);
                            }
                        }
                    },
                    {
                        className: "text-end text-nowrap",
                        name: "Actions",
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
                                                // View the item
                                                Forms.view(item);
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
                                                Forms.edit(item.Id, () => {
                                                    // Refresh the table
                                                    dashboard.refresh(DataSource.List.Items);
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