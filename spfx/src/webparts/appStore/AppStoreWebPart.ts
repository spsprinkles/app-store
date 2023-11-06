import { DisplayMode, Environment, Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneLabel, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart, WebPartContext } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme, ISemanticColors } from '@microsoft/sp-component-base';
import * as strings from 'AppStoreWebPartStrings';

export interface IAppStoreWebPartProps {
  appCatalogUrl: string;
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
    sourceUrl?: string;
  }) => void;
  setAppCatalogUrl: (url: string) => void;
  updateTheme: (currentTheme: Partial<ISemanticColors>) => void;
  version: string;
};

export default class AppStoreWebPart extends BaseClientSideWebPart<IAppStoreWebPartProps> {
  public render(): void {
    // Set the app catalog url
    AppStore.setAppCatalogUrl(this.properties.appCatalogUrl);

    // Render the application
    AppStore.render({
      el: this.domElement,
      context: this.context,
      displayMode: this.displayMode,
      envType: Environment.type
    });
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
