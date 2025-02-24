import { Navigation } from "dattatable";
import { Components } from "gd-sprest-bs";
import { generateIcon } from "gd-sprest-bs/build/icons/generate.js";
import { DataSource } from "./ds";
import { InstallationModal } from "./install";
import { LoadForm } from "./loadForm";
import Strings from "./strings";
import { Security } from "./security";
import { Tabs } from "./tabs";
import { IReportProps } from "./tabs/reports";
import { ISearchProps } from "./tabs/searchProp";

// App Properties
export interface IProp {
    description: string;
    disabled: boolean;
    label: string;
}
export interface IAppProps {
    cloudEnv?: string;
    context?: any;
    el: HTMLElement;
    maxStorageDesc?: string;
    maxStorageSize?: number;
    reportProps?: IReportProps;
    searchProps?: ISearchProps;
    siteProps: { [key: string]: IProp; }
    title?: string;
    webProps: { [key: string]: IProp; }
}

/**
 * Main Application
 */
export class App {
    private _props: IAppProps = null;

    // Constructor
    constructor(props: IAppProps) {
        this._props = props;

        // Add the class for bootstrap
        this._props.el.classList.add("bs");

        // Render the template
        this._props.el.innerHTML = `
            <div class="row">
                <div class="col-12"></div>
                <div class="col-12 mt-2"></div>
            </div>
        `;

        // Set the elements
        let elRow = this._props.el.children[0] as HTMLElement;

        // Render the dashboard
        this.renderNavigation(elRow);

        // See if data has been loaded
        if (DataSource.Site) {
            // Render the tabs
            this.renderTabs(elRow.children[1] as HTMLElement);
        } else {
            // Render the load form
            this.renderForm(elRow.children[1] as HTMLElement);
        }
    }

    // Renders the form
    private renderForm(el: HTMLElement) {
        // Render the form
        el.innerHTML = `
            <div class="row">
                <div class="col-12 my-3"></div>
                <div class="col-12 d-flex justify-content-end"></div>
            </div>
        `;

        // Render the form
        new LoadForm(el.children[0].children[0] as HTMLElement, el.children[0].children[1] as HTMLElement, () => {
            // Render the tabs
            new App(this._props);
        });
    }

    // Renders the navigation
    private renderNavigation(elRow: HTMLElement) {
        let itemsEnd: Components.INavbarItem[] = [];

        // Show the load site button if data has already been loaded
        if (DataSource.Site) {
            itemsEnd.push({
                className: "btn-outline-light",
                isButton: true,
                text: "Load Site",
                onClick: () => {
                    // Show the load form
                    LoadForm.showModal(() => {
                        // Render the tabs
                        this.renderTabs(elRow.children[1] as HTMLElement);
                    });
                }
            });
        }

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
            iconType: generateIcon(`<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 2048 2048"><path fill="var(--bs-navbar-brand-color)" d="M64 272L1216 67v1914L64 1782V272zm613 347q-54 0-98 17t-75 50-48 76-17 98q0 59 18 98t49 69 68 55 78 57q17 14 26 33t9 42q0 38-23 58t-60 20q-23 0-46-6t-44-18-40-28-32-35v173q39 26 84 40t93 14q50 0 91-15t70-45 44-70 16-92q0-62-19-104t-49-72-62-50-63-41-48-43-20-58q0-22 8-38t23-26 33-14 39-5q33 0 69 13t60 38V641q-31-11-66-16t-68-6zm731 405q0 26-10 49t-27 41-41 28-50 10V897q27 0 50 10t40 27 28 40 10 50zm616 86l-63 241-212-28q-17 26-35 50t-41 46l74 192-216 126-128-168q-60 15-123 15v-210q72 0 135-27t111-75 74-110 28-136q0-72-27-135t-75-111-110-75-136-28V461h5l79-180 241 64-27 205q27 17 52 37t48 44l188-72 125 215-167 129q13 59 13 124l187 83z"></path></svg>`, 32, 32),
            el: elRow.children[0] as HTMLElement,
            title: this._props.title || Strings.ProjectName,
            hideFilter: true,
            hideSearch: true,
            itemsEnd
        });
    }

    // Renders the tabs
    private renderTabs(el: HTMLElement) {
        // Clear the tabs element
        while (el.firstChild) { el.removeChild(el.firstChild); }

        // Render the site information
        Components.Form({
            el,
            className: "mt-1 mb-4",
            rows: [
                {
                    columns: [
                        {
                            size: 6,
                            control: {
                                label: "Top Site Url:",
                                type: Components.FormControlTypes.Readonly,
                                value: DataSource.Site.Url
                            }
                        },
                        {
                            size: 6,
                            control: {
                                label: "Sub Site Url:",
                                type: Components.FormControlTypes.Dropdown,
                                items: DataSource.SiteItems,
                                value: DataSource.Web.Id,
                                required: true,
                                onChange: item => {
                                    // Refresh the web tab
                                    tabs.refreshWebTab(item.text);
                                }
                            } as Components.IFormControlPropsDropdown
                        }
                    ]
                }
            ]
        });

        // Render the tabs
        let tabs = new Tabs(el, this._props);
    }
}