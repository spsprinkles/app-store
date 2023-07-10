import { List } from "dattatable";
import { Components, Types, Web } from "gd-sprest-bs";
import Strings from "./strings";

/**
 * App Store Item
 */
export interface IAppStoreItem extends Types.SP.ListItem {
    AdditionalInformation?: Types.SP.FieldUrlValue;
    Description: string;
    Icon: string;
    Rating?: number;
    RatingCount?: number;
    ScreenShot1: string;
    ScreenShot2?: string;
    ScreenShot3?: string;
    ScreenShot4?: string;
    ScreenShot5?: string;
    TypeOfProject: string;
    VideoURL?: Types.SP.FieldUrlValue;
}

/**
 * Rating Item
 */
export interface IRatingItem extends Types.SP.ListItem {
    AppLU?: { Id: number; Title: string; }
    Rating?: number;
}

/**
 * Data Source
 */
export class DataSource {
    // Filters
    private static _filtersTypeOfProject: Components.ICheckboxGroupItem[] = null;
    static get FiltersTypeOfProject(): Components.ICheckboxGroupItem[] { return this._filtersTypeOfProject; }
    static initFilters() {
        // Clear the filters
        this._filtersTypeOfProject = [];

        // Parse the fields
        let filterField: Types.SP.FieldChoice = null;
        for (let i = 0; i < this.List.ListFields.length; i++) {
            let fld = this.List.ListFields[i];

            // See if this is the target field
            if (fld.InternalName == "TypeOfProject") {
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
            this._filtersTypeOfProject.push({
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
                    GetAllItems: true,
                    OrderBy: ["Title"],
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