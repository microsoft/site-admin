import { Components } from "gd-sprest-bs";
import { DataSource } from "../ds";
import { IProp } from "../app";
import { IChangeRequest } from "./changes";

export interface ISearchProps {
    description: string;
    key: string;
    label: string;
    managedProperty: string;
    tabName: string;
    values: string;
}

/**
 * Search Property Tab
 */
export class SearchPropTab {
    private _currValue: string = null;
    private _el: HTMLElement = null;
    private _isRequest: boolean = null;
    private _newValue: string = null;
    private _searchProps: ISearchProps = null;

    // Constructor
    constructor(el: HTMLElement, searchProps: ISearchProps, isRequest: boolean) {
        this._el = el;
        this._searchProps = searchProps;

        // Set the current value
        this._currValue = DataSource.Site.RootWeb.AllProperties[this._searchProps.key];
        this._newValue = this._currValue;

        // See if the ability to enable custom scripts
        this._isRequest = isRequest;

        // Render the tab
        this.render();
    }

    // Gets the controls to render
    private getControls() {
        // See if the key doesn't exists
        if (this._searchProps.key == null) { return []; }

        // See if the values were provided
        if (DataSource.SearchPropItems) {
            // Return a dropdown
            return [{
                className: this._searchProps.key ? "" : "d-none",
                name: "SearchProp",
                label: this._searchProps.label || "Search Property",
                description: DataSource.SiteCustomScriptsEnabled || this._isRequest ? this._searchProps.description || "The custom property to set for search." : "<span style='color:red'>You must enable custom scripts to update this property.</span>",
                isDisabled: !(DataSource.SiteCustomScriptsEnabled || this._isRequest),
                type: Components.FormControlTypes.Dropdown,
                items: DataSource.SearchPropItems,
                value: this._currValue,
                onChange: value => {
                    // See if we are changing the value
                    if (this._currValue != value?.text) {
                        // Set the value
                        this._newValue = value?.text;
                    } else {
                        // Clear the value
                        this._newValue = null;
                    }
                }
            } as Components.IFormControlPropsDropdown];
        } else {
            // Add a textbox
            return [{
                className: this._searchProps.key ? "" : "d-none",
                name: "SearchProp",
                label: this._searchProps.label || "Search Property",
                description: this._searchProps.description || "The custom property to set for search.",
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
            } as Components.IFormControlPropsTextField];
        }
    }

    // Returns the new request items to create
    getRequests(): IChangeRequest[] {
        let requests: IChangeRequest[] = [];

        // See if we are changing the value
        if (this._currValue != this._newValue) {
            // See if we are creating a request
            if (this._isRequest) {
                // Add the request
                requests.push({
                    oldValue: this._currValue,
                    newValue: this._newValue,
                    scope: "Search",
                    request: {
                        key: "SearchProperty",
                        message: `The request set the site property to '${this._newValue}', will be processed within 5 minutes.`,
                        value: this._searchProps.key
                    },
                    url: DataSource.SiteContext.SiteFullUrl
                });
            } else {
                // Add the request
                requests.push({
                    oldValue: this._currValue,
                    newValue: this._newValue,
                    scope: "Search",
                    property: this._searchProps.key,
                    url: DataSource.SiteContext.SiteFullUrl
                });
            }
        }

        // Return the requests
        return requests;
    }

    // Renders the tab
    private render() {
        // Render the form
        Components.Form({
            el: this._el,
            controls: this.getControls()
        });
    }
}