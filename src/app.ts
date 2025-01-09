import { Navigation } from "dattatable";
import { Components } from "gd-sprest-bs";
import { DataSource } from "./ds";
import * as Forms from "./forms";
import { InstallationModal } from "./install";
import Strings from "./strings";
import { Security } from "./security";
import { Tabs } from "./tabs";

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
                <div class="col-12 mt-2"></div>
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
        new Forms.Load(this._elForm.children[0].children[0] as HTMLElement, this._elForm.children[0].children[1] as HTMLElement, () => {
            // Render the tabs
            this.renderTabs();
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
                    Forms.Load.showModal(() => {
                        // Render the tabs
                        this.renderTabs();
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
    private renderTabs() {
        // Hide the form
        this._elForm.classList.add("d-none");

        // Clear the tabs element
        while (this._elTabs.firstChild) { this._elTabs.removeChild(this._elTabs.firstChild); }

        // Render the site information
        Components.Form({
            el: this._elTabs,
            className: "mt-1 mb-2",
            rows: [
                {
                    columns: [
                        {
                            size: 6,
                            control: {
                                label: "Site Collection Url:",
                                type: Components.FormControlTypes.Readonly,
                                value: DataSource.Site.Url
                            }
                        },
                        {
                            size: 6,
                            control: {
                                label: "Site Url:",
                                type: Components.FormControlTypes.Readonly,
                                value: DataSource.Web.Url
                            }
                        }
                    ]
                }
            ]
        });

        // Render the tabs
        new Tabs(this._elTabs);
    }
}