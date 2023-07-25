import { Dashboard } from "dattatable";
import { Components, ContextInfo } from "gd-sprest-bs";
import { filterSquare } from "gd-sprest-bs/build/icons/svgs/filterSquare";
import { gearWideConnected } from "gd-sprest-bs/build/icons/svgs/gearWideConnected";
import { plusSquare } from "gd-sprest-bs/build/icons/svgs/plusSquare";
import * as jQuery from "jquery";
import * as Common from "./common";
import { DataSource, IAppStoreItem } from "./ds";
import { Forms } from "./forms";
import { InstallationModal } from "./install";
import { Security } from "./security";
import Strings from "./strings";

/**
 * Main Application
 */
export class App {
    private _dtHideCols = [4, 5, 7];
    private _dashboard: Dashboard = null;

    // Constructor
    constructor(el: HTMLElement) {
        // Render the dashboard
        this.render(el);
    }

    // Renders the dashboard
    private render(el: HTMLElement) {
        // Define the subNav items
        let subNavItems: Components.INavbarItem[] = [];
        
        // Render the Add button if IsAdmin or IsManager
        if (Security.IsAdmin || Security.IsManager) {
            subNavItems.push(
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
                                        this._dashboard.refresh(DataSource.List.Items);
                                    });
                                }
                            },
                        });
                    }
                },
            );
        }

        // Render the filters button
        subNavItems.push(
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
                                this._dashboard.showFilter();
                            }
                        },
                    });
                }
            }
        );
        
        // Create the dashboard
        this._dashboard = new Dashboard({
            el,
            hideHeader: true,
            useModal: true,
            filters: {
                items: [{
                    header: "By App Type",
                    items: DataSource.FiltersAppType,
                    onFilter: (value: string) => {
                        // Filter the table
                        this._dashboard.filter(2, value);
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
                itemsEnd: Security.IsAdmin || Security.IsManager ? [
                    {
                        className: "btn-outline-light lh-1 me-2 pt-1",
                        text: "Settings",
                        iconSize: 22,
                        iconType: gearWideConnected,
                        isButton: true,
                        items: [
                            {
                                text: "App Settings",
                                onClick: () => {
                                    // Show the install modal
                                    InstallationModal.show(true);
                                }
                            },
                            {
                                text: "List Settings",
                                onClick: () => {
                                    // Show the settings in a new tab
                                    window.open(ContextInfo.webServerRelativeUrl + "/_layouts/15/listedit.aspx?List=" + DataSource.List.ListInfo.Id);
                                }
                            }
                        ]
                    }
                ] : null,
                // Add the branding icon & text
                onRendering: (props) => {
                    // Set the class names
                    props.className = "bg-sharepoint navbar-expand rounded-top";

                    // Set the brand
                    let brand = document.createElement("div");
                    brand.className = "d-flex align-items-center";
                    brand.appendChild(Common.getIcon(36, 36, 'App Store', 'brand'));
                    brand.append(Strings.ProjectName);
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
                itemsEnd: subNavItems
            },
            table: {
                rows: DataSource.List.Items.concat(DataSource.AppCatalogItems),
                dtProps: {
                    dom: 'rt<"row"<"col-sm-4"l><"col-sm-4"i><"col-sm-4"p>>',
                    columnDefs: [
                        {
                            "targets": '_all',
                            "orderable": false,
                        },
                        {
                            "targets": [0, 8],
                            "searchable": false
                        },
                        {
                            "targets": this._dtHideCols,
                            "visible": false
                        }
                    ],
                    drawCallback: function (settings) {
                        let api = new jQuery.fn.dataTable.Api(settings) as any;
                        let div = api.table().container() as HTMLDivElement;
                        let header = api.table().header() as HTMLTableElement;
                        let table = api.table().node() as HTMLTableElement;
                        div.querySelector(".dataTables_info").classList.add("text-center");
                        div.querySelector(".dataTables_length").classList.add("pt-2");
                        div.querySelector(".dataTables_paginate").classList.add("pt-03");
                        header.classList.add("d-flex");
                        header.classList.add("d-none");
                        table.classList.add("cards");
                    },
                    lengthMenu: [5, 10, 20, 50],
                    // Order by the 1st column by default; ascending
                    order: [[1, "asc"]]
                },
                columns: [
                    {
                        className: "text-center",
                        name: "Icon",
                        title: "Icon",
                        onRenderCell: (el, column, item: IAppStoreItem) => {
                            // Clear the Icon text
                            el.innerHTML = "";
                            if (item.Icon) {
                                // Display the image
                                let img = document.createElement("img");
                                img.classList.add("icon");
                                img.src = item.Icon;
                                el.appendChild(img);
                            } else {
                                // Get the icon
                                let icon = Common.getIcon(70, 70, item.AppType);
                                el.appendChild(icon);
                            }
                        }
                    },
                    {
                        name: "Title",
                        title: "App Name",
                        onRenderCell: (el, column, item: IAppStoreItem) => {
                            el.innerHTML = `<label>${column.title}:</label>${item.Title}`;
                            el.setAttribute("data-filter", item.Title);
                        }
                    },
                    {
                        name: "AppType",
                        title: "App Type",
                        onRenderCell: (el, column, item: IAppStoreItem) => {
                            el.innerHTML = `<label>${column.title}:</label>${item.AppType}`;
                            el.setAttribute("data-filter", item.AppType);
                        }
                    },
                    {
                        name: "Description",
                        title: "Description",
                        onRenderCell: (el, column, item: IAppStoreItem) => {
                            el.innerHTML = `<label>${column.title}:</label><div class="shrink">${item.Description}</div>`;
                            el.setAttribute("data-filter", item.Description);
                        }
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
                            el.innerHTML = `<label>${column.title}:</label>${strDevs.join("<br/>")}`;
                            el.setAttribute("data-filter", strDevs.join(","));
                        }
                    },
                    {
                        name: "Organization",
                        title: "Organization",
                        onRenderCell: (el, column, item: IAppStoreItem) => {
                            el.innerHTML = `<label>${column.title}:</label>${item.Organization}`;
                            el.setAttribute("data-filter", item.Organization);
                        }
                    },
                    {
                        name: "MoreInfo",
                        title: "More Info",
                        onRenderCell: (el, column, item: IAppStoreItem) => {
                            el.innerHTML = `<label>${column.title}:</label>`;
                            el.setAttribute("data-filter", item.MoreInfo ? item.MoreInfo.Description : "");
                            // Ensure a value exists
                            if (item.MoreInfo) {
                                // Render the link
                                let elLink = document.createElement("a");
                                elLink.text = (item.MoreInfo ? item.MoreInfo.Description : "") || "Link";
                                elLink.href = (item.MoreInfo ? item.MoreInfo.Url : "") || "#";
                                elLink.target = "_blank";
                                el.appendChild(elLink);
                            } else {
                                el.innerHTML += "&nbsp;";
                            }
                        }
                    },
                    {
                        name: "SupportURL",
                        title: "Support",
                        onRenderCell: (el, column, item: IAppStoreItem) => {
                            el.innerHTML = `<label>${column.title}:</label>`;
                            el.setAttribute("data-filter", item.SupportURL ? item.SupportURL.Description : "");
                            // Ensure a value exists
                            if (item.SupportURL) {
                                // Render the link
                                let elLink = document.createElement("a");
                                elLink.text = (item.SupportURL ? item.SupportURL.Description : "") || "Link";
                                elLink.href = (item.SupportURL ? item.SupportURL.Url : "") || "#";
                                elLink.target = "_blank";
                                el.appendChild(elLink);
                            } else {
                                el.innerHTML += "&nbsp;";
                            }
                        }
                    },
                    {
                        className: "text-center",
                        name: "Actions",
                        onRenderCell: (el, column, item: IAppStoreItem) => {
                            let tooltips: Components.ITooltipProps[] = [];

                            // Add the Details button tooltip
                            tooltips.push({
                                content: "Click to view the item.",
                                btnProps: {
                                    text: "Details",
                                    type: Components.ButtonTypes.OutlinePrimary,
                                    onClick: () => {
                                        // View the item details
                                        Forms.view(item);
                                    }
                                }
                            });

                            // Add the Edit button tooltip if IsAdmin or IsManager
                            if (Security.IsAdmin || Security.IsManager) {
                                tooltips.push({
                                    content: "Click to edit the item.",
                                    btnProps: {
                                        text: "Edit",
                                        type: Components.ButtonTypes.OutlinePrimary,
                                        isDisabled: item.IsAppCatalogItem ? true : false,
                                        onClick: () => {
                                            // Edit the item
                                            Forms.edit(item.Id, () => {
                                                // Refresh the table
                                                this._dashboard.refresh(DataSource.List.Items);
                                            });
                                        }
                                    }
                                });
                            }

                            // Render the action tooltips
                            let ttg = Components.TooltipGroup({
                                el,
                                className: "shrink",
                                isSmall: true,
                                tooltips
                            });

                            // Cast the toolTipGroup element properly
                            let ttgEl = ttg.el as HTMLElement;

                            // Add click event to grow/shrink the card
                            ttgEl.addEventListener("click", (e) => {
                                // Only grow/shrink if the click is outside the button group (::after)
                                if (e.offsetX > ttgEl.offsetWidth) {
                                    let _class = 'shrink';
                                    let hide = ttgEl.classList.contains(_class);

                                    // Get the datatable object
                                    let _dt = this._dashboard.Datatable as any;
                                    let table = _dt.datatable.table();

                                    // Get the description column by id
                                    let el = table.column(3).nodes().to$() as HTMLElement;
                                    let description = el[0].lastChild as HTMLDivElement;

                                    // Update the shrink class on description & tooltip group
                                    if (hide) {
                                        description.classList.remove(_class);
                                        ttgEl.classList.remove(_class);
                                    } else {
                                        description.classList.add(_class);
                                        ttgEl.classList.add(_class);
                                    }

                                    // Show or Hide the table columns
                                    table.columns(this._dtHideCols).visible(hide);
                                }
                            });
                        }
                    }
                ]
            }
        });
    }
}