import { Components, Helper } from "gd-sprest-bs";
import * as moment from "moment";
import { IAppProps } from "../app";

/**
 * Access
 */
export class AccessTab {
    private _appProps: IAppProps = null;
    private _el: HTMLElement = null;

    // Constructor
    constructor(el: HTMLElement, appProps: IAppProps) {
        this._appProps = appProps;
        this._el = el;

        // Render the tab
        this.render();
    }

    // Returns the permissions for the admins/owners
    private getPermissions() {
        // TODO
    }

    // Renders the audit form
    private render() {
        Components.Alert({
            el: this._el,
            header: "TODO",
            content: "This section is under development."
        });
    }
}