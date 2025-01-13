import { Components } from "gd-sprest-bs";
import { DataSource, IRequest } from "../ds";

export interface ISearchProp {
    description: string;
    key: string;
    label: string;
    tabName: string;
}

/**
 * Search Property Tab
 */
export class SearchPropTab {
    private _currValue: string = null;
    private _el: HTMLElement = null;
    private _newValue: string = null;
    private _searchProp: ISearchProp = null;

    // Constructor
    constructor(el: HTMLElement, searchProp: ISearchProp) {
        this._el = el;
        this._searchProp = searchProp;

        // Set the current value
        this._currValue = DataSource.Site.RootWeb.AllProperties[this._searchProp.key];

        // Render the tab
        this.render();
    }

    // Renders the tab
    private render() {
        Components.Form({
            el: this._el,
            className: "mt-2",
            controls: [
                {
                    className: this._searchProp.key ? "" : "d-none",
                    name: "SearchProp",
                    label: this._searchProp.label || "Search Property",
                    description: this._searchProp.description || "The custom property to set for search.",
                    type: Components.FormControlTypes.TextField,
                    value: this._currValue,
                    onChange: value => {
                        // See if we are changing the value
                        if (this._currValue != value) {
                            // Set the value
                            this._newValue = value;
                        } else {
                            // Clear the value
                            this._newValue = null;
                        }
                    }
                } as Components.IFormControlPropsTextField
            ]
        });
    }
}