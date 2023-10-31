import { Modal } from "dattatable";
import { Components, ContextInfo } from "gd-sprest-bs";
import { IAppStoreItem } from "../ds";
import { Security } from "../security";
import * as Common from "../common";
import { CopyTemplate } from "./copyTemplate";
import { CreateTemplate } from "./createTemplate";

/**
 * Templates Modal
 */
export class TemplatesModal {
    // Renders the main form for the modal
    static show(appItem: IAppStoreItem, listNames: string[] = [], webUrl: string = ContextInfo.webServerRelativeUrl) {
        // Clear the modal
        Modal.clear();

        // Set the header
        Modal.setHeader("Deploy Dataset: " + appItem.Title);
        Modal.HeaderElement.prepend(Common.getIcon(28, 28, appItem.AppType + ' Template', 'icon-svg me-2'));

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
            onClick: (tab) => {
                // See if the create templates tab was clicked
                if (tab.tabName == "Copy Template" && !tab.elTab.classList.contains("disabled")) {
                    // Render the footer
                    CopyTemplate.renderFooter(Modal.FooterElement, appItem, listNames);
                } else {
                    // Render the footer
                    CreateTemplate.renderFooter(Modal.FooterElement, appItem);
                }
            },
            onRendered: () => {
                // Render the first tab footer
                CreateTemplate.renderFooter(Modal.FooterElement, appItem);
            },
            items: [
                {
                    isActive: true,
                    title: "Create Template",
                    isDisabled: !(Security.IsAdmin || Security.IsManager || isDeveloper),
                    onRenderTab: (el) => {
                        // Render the form
                        CreateTemplate.renderForm(el, webUrl);

                        // Render the footer
                        CreateTemplate.renderFooter(Modal.FooterElement, appItem);
                    }
                },
                {
                    title: "Copy Template",
                    onRenderTab: (el) => {
                        // Render the form
                        CopyTemplate.renderForm(el, appItem, listNames);

                        // Render the footer
                        CopyTemplate.renderFooter(Modal.FooterElement, appItem, listNames);
                    }
                }
            ]
        });

        // Show the modal
        Modal.show();
    }
}