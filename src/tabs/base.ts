import { IProp } from "../app";
import { DataSource, IRequest } from "../ds";
import { IChangeRequest } from "./changes";

export class Tab<IProps = { [key: string]: string | number | boolean }, IRequestItems = { [key: string]: IRequest }, INewProps = { [key: string]: string | number | boolean }> {
    protected _props: { [key: string]: IProp; } = null;
    protected _el: HTMLElement = null;

    // The current values
    protected _currValues: IProps = {} as any;

    // The new values requested
    protected _newValues: INewProps = {} as any;

    // The new request items to create
    protected _requestItems: IRequestItems = {} as any;

    // The scope of the tab
    protected _scope: "Site" | "Web" = null;

    // Constructor
    constructor(el: HTMLElement, props: { [key: string]: IProp; }, scope?: "Site" | "Web") {
        // Set the properties
        this._el = el;
        this._props = props;
        this._scope = scope;
    }

    // Returns the site properties to update
    getProps(): INewProps { return this._newValues; }

    // Returns the new request items to create
    getRequests(): IChangeRequest[] {
        let requests: IChangeRequest[] = [];
        let url = this._scope == "Site" ? DataSource.SiteContext.SiteFullUrl : DataSource.Web.Url;

        // Parse the new values
        for (let key in this._newValues) {
            // Append the request
            requests.push({
                oldValue: (this._currValues as any)[key],
                newValue: (this._newValues as any)[key],
                scope: this._scope,
                property: key,
                url
            });
        }

        // Get the request items to create
        for (let key in this._requestItems) {
            let request: IRequest = this._requestItems[key] as any;

            // Append the request
            requests.push({
                oldValue: (this._currValues as any)[key],
                newValue: request.value,
                scope: this._scope,
                request,
                url
            });
        }

        // Return the requests
        return requests.concat(this.onGetRequests());
    }

    // Event for getting custom requests
    onGetRequests(): IChangeRequest[] { return []; }
}