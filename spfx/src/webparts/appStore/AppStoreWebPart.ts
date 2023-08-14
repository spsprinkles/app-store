import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneLabel, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart, WebPartContext } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme, ISemanticColors } from '@microsoft/sp-component-base';
import * as strings from 'AppStoreWebPartStrings';

// Reference the solution
import "../../../../dist/app-store.min.js";
declare const AppStore: {
  description: string;
  render: (el: HTMLElement, context: WebPartContext) => void;
  setAppCatalogUrl: (url: string) => void;
  updateTheme: (currentTheme: Partial<ISemanticColors>) => void;
  version: string;
};

export interface IAppStoreWebPartProps {
  appCatalogUrl: string;
}

export default class AppStoreWebPart extends BaseClientSideWebPart<IAppStoreWebPartProps> {
  private _hasRendered: boolean = false;

  public render(): void {
    // See if have rendered the solution
    if (this._hasRendered) {
      // Clear the element
      while (this.domElement.firstChild) { this.domElement.removeChild(this.domElement.firstChild); }
    }

    // Set the app catalog url
    AppStore.setAppCatalogUrl(this.properties.appCatalogUrl);

    // Render the application
    AppStore.render(this.domElement, this.context);

    // Set the flag
    this._hasRendered = true;
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    // Update the theme
    AppStore.updateTheme(currentTheme.semanticColors);
  }

  protected get dataVersion(): Version {
    return Version.parse(AppStore.version);
  }


  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          groups: [
            {
              groupFields: [
                PropertyPaneTextField('appCatalogUrl', {
                  label: strings.AppCatalogUrlFieldLabel,
                  description: strings.AppCatalogUrlFieldDescription
                }),
                PropertyPaneLabel('version', {
                  text: "v" + AppStore.version
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
