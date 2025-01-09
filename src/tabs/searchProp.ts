import { Components } from "gd-sprest-bs";
import { IRequest } from "../ds";

/**
 * Search Property Tab
 */
export class SearchPropTab {
    private _el: HTMLElement = null;

    // Constructor
    constructor(el: HTMLElement) {
        this._el = el;

        // Render the tab
        this.render();
    }

    // Renders the tab
    private render() {
        this._el.innerHTML = "This feature is coming soon. This will allow you to set a custom search property on the root site.";
    }
}