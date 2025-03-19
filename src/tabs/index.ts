import { LoadingDialog } from "dattatable";
import { Components } from "gd-sprest-bs";
import { AppPermissionsTab } from "./appPermissions";
import { ChangesTab, IChangeRequest } from "./changes";
import { isEmpty } from "./common";
import { IAppProps } from "../app";
import { DataSource } from "../ds";
import { FeaturesTab } from "./features";
import { InfoTab } from "./info";
import { ManagementTab } from "./management";
import { SearchPropTab } from "./searchProp";
import { ReportsTab } from "./reports";
import { WebTab } from "./web";

/**
 * Tabs
 */
export class Tabs {
    private _el: HTMLElement = null;
    private _elWebTab: HTMLElement = null;
    private _webRequests: IChangeRequest[];
    private _tabAppPermissions: AppPermissionsTab = null;
    private _tabChanges: ChangesTab = null;
    private _tabFeatures: FeaturesTab = null;
    private _tabManagement: ManagementTab = null;
    private _tabReports: ReportsTab = null;
    private _tabSearch: SearchPropTab = null;
    private _tabWeb: WebTab = null;

    // Constructor
    constructor(el: HTMLElement, appProps: IAppProps) {
        this._el = el;
        this._webRequests = [];

        // See if the search properties exist
        if (appProps.searchProps?.values) {
            // Set the data source
            DataSource.SearchPropItems = appProps.searchProps.values;
        }

        // Render the tabs
        this.render(appProps);
    }

    // Renders the tabs
    private render(appProps: IAppProps) {
        // Set the items
        let items: Components.IListGroupItem[] = [
            {
                tabName: "Information",
                isActive: true,
                onRender: (el) => {
                    // Render the tab
                    new InfoTab(el, appProps.siteProps);
                }
            },
            {
                tabName: "Management",
                onRender: (el) => {
                    // Render the tab
                    this._tabManagement = new ManagementTab(el, appProps.siteProps, appProps.maxStorageSize, appProps.maxStorageDesc);
                }
            },
            {
                tabName: "Features",
                onRender: (el) => {
                    // Render the tab
                    this._tabFeatures = new FeaturesTab(el, appProps.siteProps);
                }
            },
            {
                tabName: "App Permissions",
                onRender: (el) => {
                    // Render the tab
                    this._tabAppPermissions = new AppPermissionsTab(el);
                }
            },
            {
                tabName: "Audit Tools",
                onRender: (el) => {
                    // Render the tab
                    this._tabReports = new ReportsTab(el, appProps.reportProps, appProps.searchProps);
                }
            }
        ];

        // See if we are customizing a search property
        if (!isEmpty(appProps.searchProps) && appProps.searchProps.key) {
            items.splice(1, 0, {
                tabName: appProps.searchProps.tabName || "Search Property",
                onRender: (el) => {
                    // Render the tab
                    this._tabSearch = new SearchPropTab(el, appProps.searchProps, appProps.siteProps["CustomScriptsEnabled"].disabled);
                }
            });
        }

        // Add the webs
        items.push({
            tabName: DataSource.Site.RootWeb.Id == DataSource.Web.Id ? "Top Site" : "Sub Site",
            onRenderTab: el => {
                // Set the tab reference
                this._elWebTab = el;
            },
            onRender: (el) => {
                // Render the tab
                this._tabWeb = new WebTab(el, appProps.webProps);
            }
        });

        // Add the changes
        items.push({
            tabName: "Changes",
            type: Components.ListGroupItemTypes.Success,
            onRender: (el) => {
                // Render the changes
                this._tabChanges = new ChangesTab(el);
            },
            onClick: () => {
                // Get the changes
                let changes = this._tabFeatures.getRequests().concat(
                    this._tabManagement.getRequests(),
                    this._webRequests,
                    this._tabWeb.getRequests()
                );

                // See if the search tab exists
                if (this._tabSearch) {
                    // Append the search request
                    changes = changes.concat(this._tabSearch.getRequests());
                }

                // Set the changes
                this._tabChanges.setChanges(changes);
            }
        });

        // Render the tabs
        Components.ListGroup({
            colWidth: 12,
            el: this._el,
            isHorizontal: true,
            isTabs: true,
            items,
            tabClassName: "mt-2"
        });
    }

    // Method to refresh the web tab
    refreshWebTab(url: string) {
        // Show a loading dialog
        LoadingDialog.setHeader("Loading Web");
        LoadingDialog.setBody("Loading the selected web...");
        LoadingDialog.show();

        // Update the requests
        this._webRequests = this._webRequests.concat(this._tabWeb.getRequests());

        // Load the web
        DataSource.loadWebInfo(url).then(() => {
            // Update the tab name
            this._elWebTab.innerHTML = DataSource.Site.RootWeb.Id == DataSource.Web.Id ? "Top Site" : "Sub Site";

            // Refresh the web tab
            this._tabWeb.refresh();

            // Hide the loading dialog
            LoadingDialog.hide();
        });
    }
}