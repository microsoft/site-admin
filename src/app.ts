import { Navigation } from "dattatable";
import { ISiteInfo } from "./ds";
import * as Forms from "./forms";
import { InstallationModal } from "./install";
import Strings from "./strings";
import { Security } from "./security";
import { Components } from "gd-sprest-bs";

// App Properties
export interface IAppProps {
    context?: any;
    el: HTMLElement;
    disableSiteProps?: string[];
    disableWebProps?: string[];
    searchProp?: Forms.ISearchProp;
}

/**
 * Main Application
 */
export class App {
    private _elNav: HTMLElement;
    private _elForm: HTMLElement;
    private _elTabs: HTMLElement;
    private _props: IAppProps = null;

    // Constructor
    constructor(props: IAppProps) {
        this._props = props;

        // Clear the element
        while (this._props.el.firstChild) { this._props.el.removeChild(this._props.el.firstChild); }

        // Add the class for bootstrap
        this._props.el.classList.add("bs");

        // Render the template
        this._props.el.innerHTML = `
            <div class="row">
                <div class="col-12"></div>
                <div class="col-12"></div>
                <div class="col-12"></div>
            </div>
        `;

        // Set the elements
        let elRow = this._props.el.children[0];
        this._elNav = elRow.children[0] as HTMLElement;
        this._elForm = elRow.children[1] as HTMLElement;
        this._elTabs = elRow.children[2] as HTMLElement;

        // Render the dashboard
        this.renderNavigation();

        // Render the form
        this.renderForm();
    }

    // Renders the form
    private renderForm() {
        // Render the form
        this._elForm.innerHTML = `
            <div class="row">
                <div class="col-12 my-3"></div>
                <div class="col-12 d-flex justify-content-end"></div>
            </div>
        `;

        // Render the form
        new Forms.Load(this._elForm.children[0].children[0] as HTMLElement, this._elForm.children[0].children[1] as HTMLElement, siteInfo => {
            // Render the tabs
            this.renderTabs(siteInfo);
        });
    }

    // Renders the navigation
    private renderNavigation() {
        // Create the items to show
        let itemsEnd = [
            {
                className: "btn-outline-light",
                isButton: true,
                text: "Load Site",
                onClick: () => {
                    // Show the load form
                    Forms.Load.showModal(siteInfo => {
                        // Render the tabs
                        this.renderTabs(siteInfo);
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

    // Renders the tabs
    private renderTabs(siteInfo: ISiteInfo) {
        // Hide the form
        this._elForm.classList.add("d-none");

        // Render the tabs
        Components.ListGroup({
            el: this._elTabs,
            isTabs: true,
            items: [
                {
                    tabName: "Site Collection",
                    isActive: true,
                    onRender: (el) => {
                        // Render the site form
                        new Forms.Site(siteInfo.site, el, this._props.disableSiteProps);
                    }
                },
                {
                    tabName: "Site",
                    onRender: (el) => {
                        // Render the web form
                        new Forms.Web(siteInfo.web, el, this._props.disableWebProps, this._props.searchProp);
                    }
                }
            ]
        });
    }
}