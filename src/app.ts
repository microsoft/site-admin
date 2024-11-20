import { Navigation } from "dattatable";
import * as Forms from "./forms";
import { InstallationModal } from "./install";
import Strings from "./strings";
import { Security } from "./security";

// App Properties
export interface IAppProps {
    context?: any;
    el: HTMLElement;
    disableSiteProps?: string[];
    disableWebProps?: string[];
    searchProp?: Forms.ISearchProp;
    siteUrls?: string[];
    webUrls?: string[];
}

/**
 * Main Application
 */
export class App {
    private _elNav: HTMLElement;
    private _elSite: HTMLElement;
    private _elWeb: HTMLElement;

    // Constructor
    constructor(props: IAppProps) {
        // Clear the element
        while (props.el.firstChild) { props.el.removeChild(props.el.firstChild); }

        // Add the class for bootstrap
        props.el.classList.add("bs");

        // Render the template
        props.el.innerHTML = `
            <div class="row">
                <div class="col-12"></div>
                <div class="col-6"></div>
                <div class="col-6"></div>
            </div>
        `;

        // Set the elements
        let elRow = props.el.children[0];
        this._elNav = elRow.children[0] as HTMLElement;
        this._elSite = elRow.children[1] as HTMLElement;
        this._elWeb = elRow.children[2] as HTMLElement;

        // Render the dashboard
        this.renderNavigation(props);
    }

    // Renders the navigation
    private renderNavigation(props: IAppProps) {
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
                        new Forms.Web(siteInfo.web, this._elWeb, props.disableWebProps, props.siteUrls, props.searchProp);
                        new Forms.Site(siteInfo.site, this._elSite, props.disableSiteProps, props.webUrls);
                    });
                }
            }
        ];

        // See if this is the admin
        if (Security.IsAdmin) {
            // Add the settings for the app
            itemsEnd.push({
                className: "btn-outline-light ms-2",
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