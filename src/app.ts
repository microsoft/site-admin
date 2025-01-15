import { Navigation } from "dattatable";
import { Components } from "gd-sprest-bs";
import { SharePoint } from "gd-sprest-bs/build/icons/custom/sharePoint";
import { DataSource } from "./ds";
import { InstallationModal } from "./install";
import { LoadForm } from "./loadForm";
import Strings from "./strings";
import { Security } from "./security";
import { Tabs } from "./tabs";
import { ISearchProp } from "./tabs/searchProp";

// App Properties
export interface IProp {
    description: string;
    disabled: boolean;
    label: string;
}
export interface IAppProps {
    context?: any;
    el: HTMLElement;
    searchProp?: ISearchProp;
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
            iconType: SharePoint,
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