import { Components } from "gd-sprest-bs";
import { IRequest } from "../ds";

/**
 * Changes Tab
 * Displays the changes and confirmation button to apply them.
 */
export class ChangesTab {
    private _el: HTMLElement = null;

    // Constructor
    constructor(el: HTMLElement) {
        this._el = el;

        // Render the tab
        this.render();
    }

    // Renders the tab
    private render() {
    }
}