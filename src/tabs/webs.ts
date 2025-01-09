import { Components } from "gd-sprest-bs";
import { ITab } from "./tab.d";
import { IRequest } from "../ds";

/**
 * Webs Tab
 * Displays the root and sub-webs of the site collection.
 */
export class WebsTab implements ITab {
    private _el: HTMLElement = null;

    // Constructor
    constructor(el: HTMLElement) {
        this._el = el;

        // Render the tab
        this.render();
    }

    getProps(): { [key: string]: string | number | boolean; } {
        let props = {};

        return props;
    }

    getRequests(): IRequest[] {
        let requests: IRequest[] = [];

        return requests;
    }

    // Renders the tab
    private render() {
    }
}