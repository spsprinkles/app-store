import { DisplayMode, Environment, Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneHorizontalRule, PropertyPaneLabel, PropertyPaneLink, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart, WebPartContext } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import * as strings from 'AppStoreWebPartStrings';

export interface IAppStoreWebPartProps {
  appCatalogUrl: string;
  title: string;
}

// Reference the solution
import "../../../../dist/app-store.min.js";
declare const AppStore: {
  description: string;
  getLogo: () => SVGImageElement;
  render: (props: {
    el: HTMLElement;
    context?: WebPartContext;
    displayMode?: DisplayMode;
    envType?: number;
    title?: string;
    sourceUrl?: string;
  }) => void;
  setAppCatalogUrl: (url: string) => void;
  title: string;
  updateTheme: (currentTheme: Partial<IReadonlyTheme>) => void;
};

export default class AppStoreWebPart extends BaseClientSideWebPart<IAppStoreWebPartProps> {
  private _hasRendered: boolean = false;

  public render(): void {
    // See if have rendered the solution
    if (this._hasRendered) {
      // Clear the element
      while (this.domElement.firstChild) { this.domElement.removeChild(this.domElement.firstChild); }
    }

    // Set the default property values
    if (!this.properties.title) { this.properties.title = AppStore.title; }
    // Set the app catalog url
    AppStore.setAppCatalogUrl(this.properties.appCatalogUrl);

    // Render the application
    AppStore.render({
      el: this.domElement,
      context: this.context,
      displayMode: this.displayMode,
      envType: Environment.type,
      title: this.properties.title
    });

    // Set the flag
    this._hasRendered = true;
  }

  // Add the logo to the PropertyPane Settings panel
  protected onPropertyPaneRendered(): void {
    const setLogo = setInterval(() => {
      let closeBtn = document.querySelectorAll("div.spPropertyPaneContainer div[aria-label='Solution Center property pane'] button[data-automation-id='propertyPaneClose']");
      if (closeBtn) {
        closeBtn.forEach((el: HTMLElement) => {
          let parent = el.parentElement;
          if (parent && !(parent.firstChild as HTMLElement).classList.contains("logo")) { parent.prepend(AppStore.getLogo()) }
        });
        clearInterval(setLogo);
      }
    }, 50);
  }
  
  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    // Update the theme
    AppStore.updateTheme(currentTheme);
  }

  protected get dataVersion(): Version {
    return Version.parse(this.context.manifest.version);
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          groups: [
            {
              groupName: "Settings:",
              groupFields: [
                PropertyPaneTextField('title', {
                  label: strings.TitleFieldLabel,
                  description: strings.TitleFieldDescription
                }),
                PropertyPaneTextField('appCatalogUrl', {
                  label: strings.AppCatalogUrlFieldLabel,
                  description: strings.AppCatalogUrlFieldDescription
                })
              ]
            }
          ]
        },
        {
          groups: [
            {
              groupName: "About this app:",
              groupFields: [
                PropertyPaneLabel('version', {
                  text: "Version: " + this.context.manifest.version
                }),
                PropertyPaneLabel('description', {
                  text: AppStore.description
                }),
                PropertyPaneLabel('about', {
                  text: "We think adding sprinkles to a donut just makes it better! SharePoint Sprinkles builds apps that are sprinkled on top of SharePoint, making your experience even better. Check out our site below to discover other SharePoint Sprinkles apps, or connect with us on GitHub."
                }),
                PropertyPaneLabel('support', {
                  text: "Are you having a problem or do you have a great idea for this app? Visit our GitHub link below to open an issue and let us know!"
                }),
                PropertyPaneHorizontalRule(),
                PropertyPaneLink('supportLink', {
                  href: "https://www.spsprinkles.com/",
                  text: "SharePoint Sprinkles",
                  target: "_blank"
                }),
                PropertyPaneLink('sourceLink', {
                  href: "https://github.com/spsprinkles/app-store/",
                  text: "View Source on GitHub",
                  target: "_blank"
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
