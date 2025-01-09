import { IRequest } from "../ds";

export class Tab<IProps = { [key: string]: string | number | boolean }, IRequestItems = { [key: string]: IRequest }, INewProps = { [key: string]: string | number | boolean }> {
    protected _disableProps: string[] = null;
    protected _el: HTMLElement = null;

    // The current values
    protected _currValues: IProps = {} as any;

    // The new values requested
    protected _newValues: INewProps = {} as any;

    // The new request items to create
    protected _requestItems: IRequestItems = {} as any;

    // Constructor
    constructor(el: HTMLElement, disableProps: string[] = []) {
        // Set the properties
        this._disableProps = disableProps;
        this._el = el;
    }

    // Returns the site properties to update
    getProps(): INewProps { return this._newValues; }

    // Returns the new request items to create
    getRequests(): IRequest[] {
        // Get the request items to create
        let requests: IRequest[] = [];
        for (let key in this._requestItems) {
            // Append the request
            requests.push(this._requestItems[key] as IRequest);
        }

        // Return the requests
        return requests;
    }
}