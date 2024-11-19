import { Navigation } from "dattatable";
import * as Forms from "./forms";
import { InstallationModal } from "./install";
import Strings from "./strings";
import { Security } from "./security";

/**
 * Main Application
 */
export class App {
    private _elNav: HTMLElement;
    private _elSite: HTMLElement;
    private _elWeb: HTMLElement;

    // Constructor
    constructor(el: HTMLElement, hideProps: string[] = []) {
        // Clear the element
        while (el.firstChild) { el.removeChild(el.firstChild); }

        // Add the class for bootstrap
        el.classList.add("bs");

        // Render the template
        el.innerHTML = `
            <div class="row">
                <div class="col-12"></div>
                <div class="col-6"></div>
                <div class="col-6"></div>
            </div>
        `;

        // Set the elements
        let elRow = el.children[0];
        this._elNav = elRow.children[0] as HTMLElement;
        this._elSite = elRow.children[1] as HTMLElement;
        this._elWeb = elRow.children[2] as HTMLElement;

        // Render the dashboard
        this.renderNavigation(hideProps);
    }

    // Renders the navigation
    private renderNavigation(hideProps: string[]) {
        // Create the items to show
        let itemsEnd = [
            {
                className: "btn-outline-light",
                isButton: true,
                text: "Load Site",
                onClick: () => {
                    // Show the load form
                    Forms.Load.show((siteInfo) => {
                        // Render the forms for this site
                        new Forms.Web(siteInfo.web, this._elWeb, hideProps);
                        new Forms.Site(siteInfo.site, this._elSite, hideProps);
                    });
                }
            }
        ];

        // See if this is the admin
        if (Security.IsAdmin) {
            // Add the settings for the app
            itemsEnd.push({
                className: "btn-outline-light me-2",
                isButton: true,
                text: "Settings",
                onClick: () => {
                    // Show the app settings
                    InstallationModal.show(true);
                }
            });
        }

        // Render the navigation
        new Navigation({
            el: this._elNav,
            title: Strings.ProjectName,
            hideFilter: true,
            hideSearch: true,
            itemsEnd
        });
    }
}