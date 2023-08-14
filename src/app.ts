import { Dashboard } from "dattatable";
import { Components, ContextInfo } from "gd-sprest-bs";
import { filterSquare } from "gd-sprest-bs/build/icons/svgs/filterSquare";
import { gearWideConnected } from "gd-sprest-bs/build/icons/svgs/gearWideConnected";
import { pencilSquare } from "gd-sprest-bs/build/icons/svgs/pencilSquare";
import { plusSquare } from "gd-sprest-bs/build/icons/svgs/plusSquare";
import { window_ } from "gd-sprest-bs/build/icons/svgs/window_";
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
    private _dashboard: Dashboard = null;

    // Constructor
    constructor(el: HTMLElement) {
        // Render the dashboard
        this.render(el);
    }

    // Renders the dashboard
    private render(el: HTMLElement) {
        // Define the nav items
        let navItems: Components.INavbarItem[] = [];
        let subNavItems: Components.INavbarItem[] = [];

        if (DataSource.AppCatalogUrl) {
            // Render the App Catalog button
            navItems.push(
                {
                    text: "App Catalog Dashboard",
                    onRender: (el, item) => {
                        // Clear the existing button
                        el.innerHTML = "";
                        
                        // Render a tooltip
                        Components.Tooltip({
                            el,
                            content: item.text,
                            type: Components.TooltipTypes.LightBorder,
                            btnProps: {
                                // Render the icon button
                                className: "p-1 pe-2 me-2",
                                iconClassName: "me-1",
                                iconType: Common.getIcon(25, 25, "App Dashboard", "brand"),
                                text: "App Dashboard",
                                type: Components.ButtonTypes.OutlineLight,
                                onClick: () => {
                                    window.open(DataSource.AppCatalogUrl, "_blank");
                                }
                            }
                        });
                    }
                }
            );
        }

        if (Security.IsAdmin || Security.IsManager) {
            // Render the settings menu
            navItems.push(
                {
                    className: "btn-outline-light lh-1 me-2 py-1",
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
            );

            // Render the Add button
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
                                        this._dashboard.refresh(DataSource.AppItems);
                                    });
                                }
                            }
                        });
                    }
                }
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
                        }
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
                itemsEnd: navItems,
                // Add the branding icon & text
                onRendering: (props) => {
                    // Set the class names
                    props.className = "bg-sharepoint navbar-expand rounded-top";

                    // Set the brand
                    let brand = document.createElement("div");
                    brand.className = "d-flex align-items-center";
                    brand.appendChild(Common.getIcon(36, 36, 'App Store', 'brand me-2'));
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
                rows: DataSource.AppItems,
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
                        className: "d-none",
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
                        className: "d-none",
                        onRenderCell: (el, column, item: IAppStoreItem) => {
                            el.innerHTML = `<label>${column.title}:</label>${item.Organization || ""}`;
                            el.setAttribute("data-filter", item.Organization || "");
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
                        className: "d-none",
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
                            let root = document.querySelector(':root') as HTMLElement;
                            let tooltips: Components.ITooltipProps[] = [];

                            // Add the Details button tooltip
                            tooltips.push({
                                content: "View more details",
                                btnProps: {
                                    className: "p-1 pe-2",
                                    iconClassName: "me-1",
                                    iconType: window_,
                                    iconSize: 24,
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
                                    content: "Edit the item",
                                    btnProps: {
                                        className: "p-1 pe-2",
                                        iconClassName: "me-1",
                                        iconType: pencilSquare,
                                        iconSize: 24,
                                        text: "Edit",
                                        type: Components.ButtonTypes.OutlinePrimary,
                                        isDisabled: item.IsAppCatalogItem ? true : false,
                                        onClick: () => {
                                            // Edit the item
                                            Forms.edit(item.Id, () => {
                                                // Refresh the table
                                                this._dashboard.refresh(DataSource.AppItems);
                                            });
                                        }
                                    }
                                });
                                root.style.setProperty('--shrink-right', '-4rem');
                            }

                            // Render the action tooltips
                            let ttg = Components.TooltipGroup({
                                el,
                                className: "shrink",
                                isSmall: true,
                                tooltips
                            });

                            // Add click event to grow/shrink the card
                            ttg.el.addEventListener("click", (e) => {
                                // Only grow/shrink if the click is outside the button group [::after]
                                if (e.offsetX > ttg.el.offsetWidth) {
                                    let _class = 'shrink';
                                    let hide = ttg.el.classList.contains(_class);
                                    let tr = ttg.el.closest("tr");

                                    if (hide) {
                                        // Remove the shrink class on description inner div
                                        jQuery("td:nth-child(4) :last-child", tr).removeClass(_class);
                                        // Show columns [nth-child() is not 0 index based]
                                        jQuery("td:nth-child(5), td:nth-child(6), td:nth-child(8)", tr).removeClass("d-none");
                                        // Remove the shrink class on tooltip group
                                        ttg.el.classList.remove(_class);
                                    } else {
                                        // Add the shrink class on description inner div
                                        jQuery("td:nth-child(4) :last-child", tr).addClass(_class);
                                        // Hide columns [nth-child() is not 0 index based]
                                        jQuery("td:nth-child(5), td:nth-child(6), td:nth-child(8)", tr).addClass("d-none");
                                        // Add the shrink class on tooltip group
                                        ttg.el.classList.add(_class);
                                    }
                                }
                            });
                        }
                    }
                ]
            }
        });
    }
}