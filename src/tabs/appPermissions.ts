/**
 * Application Permissions Tab
 */
export class AppPermissionsTab {
    private _el: HTMLElement = null;

    // Constructor
    constructor(el: HTMLElement) {
        this._el = el;

        // Render the tab
        this.render();
    }

    // Renders the tab
    private render() {
        this._el.classList.add("mt-2");
        this._el.innerHTML = "This feature is coming soon. This will allow you to set permissions for apps accessing data from the site collection.";
    }
}