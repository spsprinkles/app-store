import { List } from "dattatable";
import { Components, Types } from "gd-sprest-bs";
import Strings from "./strings";

/**
 * App Store Item
 */
export interface IAppStoreItem extends Types.SP.ListItem {
    AppType: string;
    Description: string;
    Developers: { results: { Id: number; EMail: string; Title: string }[] };
    Icon: string;
    Modified: string;
    MoreInfo?: Types.SP.FieldUrlValue;
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
                    Expand: ["Developers"],
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
            this.initList()
        ]);
    }
}