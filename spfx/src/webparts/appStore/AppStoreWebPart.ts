import { DisplayMode, Environment, Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneLabel, PropertyPaneTextField } from '@microsoft/sp-property-pane';
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
              groupFields: [
                PropertyPaneTextField('title', {
                  label: strings.TitleFieldLabel,
                  description: strings.TitleFieldDescription
                }),
                PropertyPaneTextField('appCatalogUrl', {
                  label: strings.AppCatalogUrlFieldLabel,
                  description: strings.AppCatalogUrlFieldDescription
                }),
                PropertyPaneLabel('version', {
                  text: "v" + this.context.manifest.version
                })
              ]
            }
          ],
          header: {
            description: AppStore.description
          }
        }
      ]
    };
  }
}
