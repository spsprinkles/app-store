import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart, WebPartContext } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme, ISemanticColors } from '@microsoft/sp-component-base';

// Reference the solution
import "../../../../dist/app-store.min.js";
declare const AppStore: {
  render: (el: HTMLElement, context: WebPartContext) => void;
  updateTheme: (currentTheme: Partial<ISemanticColors>) => void;
};

export interface IAppStoreWebPartProps { }

export default class AppStoreWebPart extends BaseClientSideWebPart<IAppStoreWebPartProps> {

  public render(): void {
    // Render the application
    AppStore.render(this.domElement, this.context);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    // Update the theme
    AppStore.updateTheme(currentTheme.semanticColors);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }
}
