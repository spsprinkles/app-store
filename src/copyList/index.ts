import { Modal } from "dattatable";
import { Components, ContextInfo } from "gd-sprest-bs";
import { IAppStoreItem } from "../ds";
import { Security } from "../security";
import { CopyLists } from "./copyLists";
import { CreateLists } from "./createLists";

/**
 * Copy List Modal
 */
export class CopyListModal {
    // Renders the main form for the modal
    static show(appItem: IAppStoreItem, listNames: string[] = [], webUrl: string = ContextInfo.webServerRelativeUrl) {
        // Clear the modal
        Modal.clear();

        // Set the header
        Modal.setHeader("Copy List");

        // Determine if this is one of the developers
        let isDeveloper = false;
        if (appItem.Developers && appItem.Developers.results.length > 0) {
            for (let i = 0; i < appItem.Developers.results.length; i++) {
                // See if this is one of the developers
                if (appItem.Developers.results[i].Id == ContextInfo.userId) {
                    // Set the flag
                    isDeveloper = true;
                    break;
                }
            }
        }

        // Renders the tabs
        Components.Nav({
            el: Modal.BodyElement,
            isTabs: true,
            isUnderline: true,
            onClick: (tab) => {
                // See if the copy lists tab was clicked
                if (tab.tabName == "Add Template" && !tab.elTab.classList.contains("disabled")) {
                    // Render the footer
                    CopyLists.renderFooter(Modal.FooterElement, appItem, listNames);
                } else {
                    // Render the footer
                    CreateLists.renderFooter(Modal.FooterElement, appItem);
                }
            },
            items: [
                {
                    isActive: true,
                    title: "Create Lists",
                    onRenderTab: (el) => {
                        // Render the form
                        CreateLists.renderForm(el, webUrl);

                        // Render the footer
                        CreateLists.renderFooter(Modal.FooterElement, appItem);
                    },
                },
                {
                    title: "Add Template",
                    isDisabled: !(Security.IsAdmin || Security.IsManager || isDeveloper),
                    onRenderTab: (el) => {
                        // Render the form
                        CopyLists.renderForm(el, appItem, listNames);

                        // Render the footer
                        CopyLists.renderFooter(Modal.FooterElement, appItem, listNames);
                    }
                }
            ]
        });

        // Show the modal
        Modal.show();
    }
}