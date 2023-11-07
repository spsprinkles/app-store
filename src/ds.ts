import { List } from "dattatable";
import { Components, ContextInfo, Types, Web } from "gd-sprest-bs";
import { Security } from "./security";
import Strings from "./strings";

/**
 * App Store Item
 */
export interface IAppStoreItem extends Types.SP.ListItemOData {
    AppType: string;
    AssociatedLists: string;
    Description: string;
    Developers: { results: { Id: number; EMail: string; Title: string }[] };
    Icon: string;
    Modified: string;
    MoreInfo?: Types.SP.FieldUrlValue;
    IsAppCatalogItem?: boolean;
    Organization: string;
    Rating?: number;
    RatingCount?: number;
    ScreenShot1: string;
    ScreenShot2?: string;
    ScreenShot3?: string;
    ScreenShot4?: string;
    ScreenShot5?: string;
    Status: string;
    SupportURL?: Types.SP.FieldUrlValue;
    VideoURL?: Types.SP.FieldUrlValue;
}

/**
 * Rating Item
 */
export interface IRatingItem extends Types.SP.ListItem {
    AppLU?: { Id: number; Title: string; }
    Comment?: string;
    Rating?: number;
}

/**
 * Data Source
 */
export class DataSource {
    // App Catalog List
    private static _appCatalogUrl: string = null;
    static get AppCatalogUrl(): string { return this._appCatalogUrl; }
    static set AppCatalogUrl(value: string) { this._appCatalogUrl = value; }
    private static _appCatalogItems: IAppStoreItem[] = null;
    static get AppCatalogItems(): IAppStoreItem[] { return this._appCatalogItems; }
    private static loadAppCatalog(): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve) => {
            // Clear the items
            this._appCatalogItems = [];

            // See if the app catalog url exist
            if (this._appCatalogUrl) {
                Web.getWebUrlFromPageUrl(this._appCatalogUrl).execute(webUrl => {
                    // Load the list information
                    Web(webUrl.GetWebUrlFromPageUrl).Lists("Developer Apps").Items().query({
                        Expand: ["AppDevelopers"],
                        Filter: "ContentType eq 'App' and (AppStatus eq 'Approved' or AppIsTenantDeployed eq 1)",
                        GetAllItems: true,
                        Select: ["*", "AppDevelopers/Title"],
                        Top: 5000
                    }).execute(items => {
                        // Parse the items
                        for (let i = 0; i < items.results.length; i++) {
                            let item: any = items.results[i];

                            // Add the item
                            this._appCatalogItems.push({
                                AppType: "SharePoint",
                                Description: item.AppDescription,
                                Developers: item.AppDevelopers,
                                Icon: item.AppThumbnailURLBase64,
                                Id: item.Id,
                                IsAppCatalogItem: true,
                                Modified: item.Modified,
                                MoreInfo: {
                                    Description: item.AppSourceControl ? item.AppSourceControl.Description : "",
                                    Url: item.AppSourceControl ? item.AppSourceControl.Url : ""
                                },
                                Organization: item.AppPublisher,
                                Rating: 0,
                                RatingCount: 0,
                                ScreenShot1: item.AppImageURL1Base64,
                                ScreenShot2: item.AppImageURL2Base64,
                                ScreenShot3: item.AppImageURL3Base64,
                                ScreenShot4: item.AppImageURL4Base64,
                                ScreenShot5: item.AppImageURL5Base64,
                                Status: item.AppStatus,
                                SupportURL: {
                                    Description: item.AppSupportURL ? item.AppSupportURL.Description : "",
                                    Url: item.AppSupportURL ? item.AppSupportURL.Url : ""
                                },
                                Title: item.Title,
                                VideoURL: {
                                    Description: item.AppVideoURL ? item.AppVideoURL.Description : "",
                                    Url: item.AppVideoURL ? item.AppVideoURL.Url : ""
                                }
                            } as any)
                        }

                        // Resolve the request
                        resolve();
                    }, () => {
                        // Resolve the request
                        resolve();
                    });
                });
            } else {
                // Resolve the request
                resolve();
            }
        });
    }

    // Reference to the app items
    // This will include both the app catalog and app store items
    static get AppItems(): IAppStoreItem[] { return this.List.Items.concat(this.AppCatalogItems); }

    // Filters
    private static _filtersAppType: Components.ICheckboxGroupItem[] = null;
    static get FiltersAppType(): Components.ICheckboxGroupItem[] { return this._filtersAppType; }
    static initFilters() {
        // Clear the filters
        this._filtersAppType = [];

        // Parse the fields
        let filterField: Types.SP.FieldChoice = null;
        for (let i = 0; i < this.List.ListFields.length; i++) {
            let fld = this.List.ListFields[i];

            // See if this is the target field
            if (fld.InternalName == "AppType") {
                // Set the field
                filterField = fld;
                break;
            }
        }

        // See if the field doesn't exists
        if (filterField == null) { return; }

        // Parse the choices
        for (let i = 0; i < filterField.Choices.results.length; i++) {
            // Add an item
            this._filtersAppType.push({
                label: filterField.Choices.results[i],
                type: Components.CheckboxGroupTypes.Switch
            });
        }
    }

    // List
    private static _list: List<IAppStoreItem> = null;
    static get List(): List<IAppStoreItem> { return this._list; }
    private static initList(): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve, reject) => {
            this._list = new List({
                listName: Strings.Lists.Main,
                itemQuery: {
                    Expand: ["Developers", "AttachmentFiles"],
                    GetAllItems: true,
                    OrderBy: ["Title"],
                    Select: ["*", "Developers/Id", "Developers/EMail", "Developers/Title",],
                    Top: 5000
                },
                onInitError: reject,
                onInitialized: () => {
                    // Initialize the filters
                    this.initFilters();

                    // Resolve the request
                    resolve();
                }
            });
        });
    }

    // Gets the item id from the query string
    static getItemIdFromQS() {
        // Get the id from the querystring
        let qs = document.location.search.split('?');
        qs = qs.length > 1 ? qs[1].split('&') : [];
        for (let i = 0; i < qs.length; i++) {
            let qsItem = qs[i].split('=');
            let key = qsItem[0];
            let value = qsItem[1];

            // See if this is the "id" key
            if (key == "ID") {
                // Return the item
                return parseInt(value);
            }
        }
    }

    // Initializes the application
    static init(): PromiseLike<any> {
        // Execute the required initialization methods
        return Promise.all([
            this.initList(),
            this.initRatingsList(),
            this.loadAppCatalog(),
            Security.init()
        ]);
    }

    // Ratings List
    private static _ratingsList: List<IRatingItem> = null;
    static get RatingsList(): List<IRatingItem> { return this._ratingsList; }
    private static initRatingsList(): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve, reject) => {
            this._ratingsList = new List({
                listName: Strings.Lists.Ratings,
                onInitError: reject,
                onInitialized: resolve
            });
        });
    }

    // Theme information
    private static _themeInfo = null;
    static get ThemeInfo() { return this._themeInfo; }
    static set ThemeInfo(value) { this._themeInfo = value; }
    static getThemeColor(name: string) { return ContextInfo.theme.accent ? ContextInfo.theme[name] : this._themeInfo[name]; }
}