import { Dashboard } from "dattatable";
import { Components, ContextInfo, ThemeManager, Types } from "gd-sprest-bs";
import { filterSquare } from "gd-sprest-bs/build/icons/svgs/filterSquare";
import { gearWideConnected } from "gd-sprest-bs/build/icons/svgs/gearWideConnected";
import { plusSquare } from "gd-sprest-bs/build/icons/svgs/plusSquare";
import * as jQuery from "jquery";
import * as Common from "./common";
import { CopyTemplate } from "./copyTemplate";
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
        // Clear the element
        while (el.firstChild) { el.removeChild(el.firstChild); }

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
                                className: "p-1 px-2 me-2",
                                iconType: Common.getIcon(22, 22, 'App Dashboard', 'brand'),
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
                    className: "btn-icon btn-outline-light me-2 p-2 py-1",
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
                            text: "Main List",
                            onClick: () => {
                                // Show the settings in a new tab
                                window.open(ContextInfo.webServerRelativeUrl + "/_layouts/15/listedit.aspx?List=" + DataSource.List.ListInfo.Id);
                            }
                        },
                        {
                            text: "Ratings List",
                            onClick: () => {
                                // Show the settings in a new tab
                                window.open(ContextInfo.webServerRelativeUrl + "/_layouts/15/listedit.aspx?List=" + DataSource.RatingsList.ListInfo.Id);
                            }
                        },
                        {
                            text: "Requests List",
                            onClick: () => {
                                // Show the settings in a new tab
                                window.open(ContextInfo.webServerRelativeUrl + "/_layouts/15/listedit.aspx?List=" + DataSource.RequestsList.ListInfo.Id);
                            }
                        },
                        {
                            text: "Developer's Group",
                            onClick: () => {
                                // Show the settings in a new tab
                                window.open(ContextInfo.webServerRelativeUrl + "/_layouts/15/people.aspx?MembershipGroupId=" + Security.DeveloperGroup.Id);
                            }
                        },
                        {
                            text: "Manager's Group",
                            onClick: () => {
                                // Show the settings in a new tab
                                window.open(ContextInfo.webServerRelativeUrl + "/_layouts/15/people.aspx?MembershipGroupId=" + Security.ManagerGroup.Id);
                            }
                        },
                        {
                            text: "Owners's Group",
                            onClick: () => {
                                // Show the settings in a new tab
                                window.open(ContextInfo.webServerRelativeUrl + "/_layouts/15/people.aspx?MembershipGroupId=" + Security.AdminGroup.Id);
                            }
                        }
                    ]
                }
            );

            // Render the Add button
            subNavItems.push(
                {
                    text: "Adds an item to the " + Strings.ProjectName,
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

        // Render the Add button
        subNavItems.push(
            {
                text: "Request an App to be added to the " + Strings.ProjectName,
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
                            iconType: Common.getIcon(24, 24, 'AppIconDefaultAdd', 'icon-svg me-1'),
                            isSmall: true,
                            text: "App Request",
                            type: Components.ButtonTypes.OutlineSecondary,
                            onClick: () => {
                                // Show the new request form
                                Forms.newRequest(() => {
                                    // Refresh the table
                                    this._dashboard.refresh(DataSource.AppItems);
                                });
                            }
                        }
                    });
                }
            }
        );

        // Render the Add button
        subNavItems.push(
            {
                text: "View app requests for the " + Strings.ProjectName,
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
                            iconType: Common.getIcon(24, 24, 'AppIconDefaultList', 'icon-svg me-1'),
                            isSmall: true,
                            text: "View Requests",
                            type: Components.ButtonTypes.OutlineSecondary,
                            onClick: () => {
                                // Show the requests modal
                                Forms.viewRequests();
                            }
                        }
                    });
                }
            }
        );

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
                        className: "p-0 pe-none text-body",
                        text: "v" + Strings.Version,
                        onRender: (el) => {
                            // Hide version footer in a modern page
                            Strings.IsClassic ? null : el.classList.add("d-none");
                        }
                    }
                ]
            },
            navigation: {
                itemsEnd: navItems,
                // Add the branding icon & text
                onRendering: (props) => {
                    // Set the class names
                    props.className = "navbar-expand rounded-top";
                    props.type = Components.NavbarTypes.Primary

                    // Set the brand
                    let brand = document.createElement("div");
                    brand.className = "d-flex align-items-center mb-1";
                    brand.appendChild(Common.getIcon(44, 44, 'App Store', 'brand'));
                    brand.append(Strings.ProjectName);
                    props.brand = brand;
                },
                // Adjust the brand alignment
                onRendered: (el) => {
                    el.querySelector("nav div.container-fluid a.navbar-brand").classList.add("p-0");
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
                            "targets": [0, 8, 9],
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
                    order: [[1, "asc"]],
                    pageLength: 10
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
                                ThemeManager.IsInverted ? img.classList.add("invert") : null;
                                img.src = item.Icon;
                                el.appendChild(img);
                            } else {
                                // Get the icon
                                let icon = Common.getIcon(70, 70, item.AppType, 'icon-type');
                                el.appendChild(icon);
                            }
                        }
                    },
                    {
                        name: "Title",
                        title: "App Name",
                        onRenderCell: (el, column, item: IAppStoreItem) => {
                            el.innerHTML = `<label>${column.title}:</label><div class="shrink title">${item.Title}</div>`;
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
                            if (item.IsAppCatalogItem) {
                                // Set the more info link
                                el.innerHTML += `<a href="${DataSource.AppCatalogUrl}?app-id=${item.Id}" target="_blank">App Dashboard (Details)</a>`;
                            } else {
                                // Render the link
                                item.MoreInfo ? el.innerHTML += `<a href="${item.MoreInfo.Url}" class="line-limit-1" target="_blank">${item.MoreInfo.Description}</a>` : el.innerHTML += "&nbsp;";
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
                            item.SupportURL ? el.innerHTML += `<a href="${item.SupportURL.Url}" class="line-limit-1" target="_blank">${item.SupportURL.Description}</a>` : el.innerHTML += "&nbsp;";
                        }
                    },
                    {
                        name: "Deploy Dataset",
                        title: "Template",
                        className: "d-none",
                        onRenderCell: (el, column, item: IAppStoreItem) => {
                            el.innerHTML = `<label>${column.title}:</label>`;

                            // See if this item has a SharePoint list associated with it
                            if (item.AssociatedLists) {
                                // Add a template button
                                Components.Tooltip({
                                    el,
                                    content: "Deploy the dataset for this solution",
                                    btnProps: {
                                        className: "p-1 pe-2",
                                        iconType: Common.getIcon(22, 22, item.AppType + (item.AppType.startsWith('Power ') ? ' ' + column.title : ''), 'icon-svg me-1'),
                                        isDisabled: !(item.AssociatedLists),
                                        isSmall: true,
                                        text: column.name,
                                        type: Components.ButtonTypes.OutlinePrimary,
                                        onClick: () => {
                                            // Get the list templates associated w/ this item
                                            let listNames = (item.AssociatedLists || "").trim().split('\n');

                                            // Display the copy list modal
                                            CopyTemplate.renderModal(item, listNames);
                                        }
                                    }
                                });

                                // Ensure this column is displayed
                                el.classList.remove("d-none");
                            }

                            // See if this is a flow and the user can download
                            let packages: Types.SP.Attachment[] = [];
                            if (item.AppType == "Power Automate" && !Strings.IsFlow3) {
                                if (item.AttachmentFiles?.results) {
                                    let promises = [];
                                    for (let i = 0; i < item.AttachmentFiles.results.length; i++) {
                                        // See if this is a flow package
                                        let promise = Forms.isFlowPackage(item.AttachmentFiles.results[i]);
                                        promise.then((result) => {
                                            if (result.isFlow) {
                                                packages.push(result.attachment);
                                            }
                                        });
                                        promises.push(promise);
                                    }

                                    // Wait for the promises to complete
                                    Promise.all(promises).then(() => {
                                        // See if a package exists
                                        if (packages.length > 0) {
                                            let showColumn = false;

                                            // Ensure this is the owner of the app or an admin
                                            if (Security.IsAdmin || Security.IsManager || Security.IsDeveloper) {
                                                // Add a button to customize the package
                                                Components.Tooltip({
                                                    el,
                                                    content: "Click to allow users to generate a custom package.",
                                                    btnProps: {
                                                        className: "p-1 pe-2",
                                                        iconType: Common.getIcon(24, 24, 'WindowEdit', 'icon-svg me-1'),
                                                        text: "Customize Package",
                                                        type: Components.ButtonTypes.OutlinePrimary,
                                                        onClick: () => {
                                                            // Show the form to customize the package
                                                            Forms.customizeFlowPackage(item, packages[0], () => {
                                                                // Refresh the table
                                                                this._dashboard.refresh(DataSource.AppItems);
                                                            });
                                                        }
                                                    }
                                                });

                                                // Set the flag
                                                showColumn = true;
                                            }

                                            // See if flow data exists
                                            if (item.FlowData) {
                                                // Add a button to generate a package
                                                Components.Tooltip({
                                                    el,
                                                    content: "Click to generate a custom package.",
                                                    btnProps: {
                                                        className: "p-1 pe-2",
                                                        iconType: Common.getIcon(22, 22, item.AppType + ' ' + column.title, 'icon-svg me-1'),
                                                        text: "Generate Package",
                                                        type: Components.ButtonTypes.OutlinePrimary,
                                                        onClick: () => {
                                                            // Show the form to generate the package
                                                            Forms.generateFlow(item, packages[0]);
                                                        }
                                                    }
                                                });

                                                // Set the flag
                                                showColumn = true;
                                            }

                                            // Show this column based on the flag
                                            showColumn ? el.classList.remove("d-none") : null;
                                        }
                                    });
                                }
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
                                    iconType: Common.getIcon(24, 24, 'AppsContent', 'icon-svg img-flip-x me-1'),
                                    text: "Details",
                                    type: Components.ButtonTypes.OutlinePrimary,
                                    onClick: () => {
                                        // View the item details
                                        Forms.view(item);
                                    }
                                }
                            });

                            // Add the Edit button tooltip if IsAdmin or IsManager or isDeveloper
                            if (Security.IsAdmin || Security.IsManager || Security.IsDeveloper) {
                                // Add the edit button
                                tooltips.push({
                                    content: "Edit the item",
                                    btnProps: {
                                        className: "p-1 pe-2",
                                        iconType: Common.getIcon(24, 24, 'WindowEdit', 'icon-svg me-1'),
                                        text: "Edit",
                                        type: Components.ButtonTypes.OutlinePrimary,
                                        isDisabled: item.IsAppCatalogItem ? true : false,
                                        onClick: () => {
                                            // Edit the item
                                            Forms.edit(item, () => {
                                                // Refresh the table
                                                this._dashboard.refresh(DataSource.AppItems);
                                            });
                                        }
                                    }
                                });
                                root.style.setProperty('--shrink-right', '-4rem');
                            }

                            // Define shrink class
                            let _class = 'shrink';

                            // Render the action tooltips
                            let ttg = Components.TooltipGroup({
                                el,
                                className: _class,
                                isSmall: true,
                                tooltips
                            });

                            // Invert the colors for grow/shrink if ThemeInfo.isInverted
                            ThemeManager.IsInverted ? ttg.el.classList.add("invert") : null;

                            // Add data attribute for Power Platform items
                            item.AppType.startsWith('Power ') ? ttg.el.setAttribute("data-ispowerplatform", "") : null;

                            // Add click event to grow/shrink the card
                            ttg.el.addEventListener("click", (e) => {
                                // Only grow/shrink if the click is outside the button group [::after]
                                if (e.offsetX > ttg.el.offsetWidth) {
                                    let hide = ttg.el.classList.contains(_class);
                                    let tdHide = "td:nth-child(5), td:nth-child(6), td:nth-child(8)";
                                    // Only show Templates for Power Platform items
                                    ttg.el.dataset["ispowerplatform"] != undefined ? tdHide += ", td:nth-child(9)" : null;
                                    let tr = ttg.el.closest("tr");

                                    if (hide) {
                                        // Remove the shrink class on title & description inner div
                                        jQuery("td:nth-child(2) :last-child, td:nth-child(4) :last-child", tr).removeClass(_class);
                                        // Show columns [nth-child() is not 0 index based]
                                        jQuery(tdHide, tr).removeClass("d-none");
                                        // Remove the shrink class on tooltip group
                                        ttg.el.classList.remove(_class);
                                    } else {
                                        // Add the shrink class on title & description inner div
                                        jQuery("td:nth-child(2) :last-child, td:nth-child(4) :last-child", tr).addClass(_class);
                                        // Hide columns [nth-child() is not 0 index based]
                                        jQuery(tdHide, tr).addClass("d-none");
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