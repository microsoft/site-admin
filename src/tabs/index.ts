import { LoadingDialog } from "dattatable";
import { Components } from "gd-sprest-bs";
import { AppPermissionsTab } from "./appPermissions";
import { ChangesTab, IChangeRequest } from "./changes";
import { isEmpty } from "./common";
import { IAppProps } from "../app";
import { DataSource } from "../ds";
import { FeaturesTab } from "./features";
import { InfoTab } from "./info";
import { ListsTab } from "./lists";
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
    private _loadOneDrive: boolean = false;
    private _webRequests: IChangeRequest[];
    private _tabAppPermissions: AppPermissionsTab = null;
    private _tabChanges: ChangesTab = null;
    private _tabFeatures: FeaturesTab = null;
    private _tabLists: ListsTab = null;
    private _tabManagement: ManagementTab = null;
    private _tabReports: ReportsTab = null;
    private _tabSearch: SearchPropTab = null;
    private _tabWeb: WebTab = null;

    // Constructor
    constructor(el: HTMLElement, appProps: IAppProps, loadOneDrive: boolean) {
        this._el = el;
        this._loadOneDrive = loadOneDrive;
        this._webRequests = [];

        // See if the search properties exist
        if (appProps.searchProps?.values) {
            // Set the data source
            DataSource.SearchPropItems = appProps.searchProps.values;
        }

        // Render the tabs
        this.render(appProps, loadOneDrive);
    }

    // Event when the webs are loaded
    onWebsLoaded(webs: Components.IDropdownItem[]) {
        // Update the webs
        this._tabReports.setWebs(webs);
    }

    // Renders the tabs
    private render(appProps: IAppProps, loadOneDrive: boolean) {
        let auditOnly = !DataSource.IsAdmin || appProps.auditOnly;

        // Show the info tab by default always
        let items: Components.IListGroupItem[] = loadOneDrive ? [] : [{
            tabName: "Information",
            isActive: true,
            onRender: (el) => {
                // Render the tab
                new InfoTab(el, appProps.siteAttestation, appProps.siteProps);
            }
        }];

        // Add the tabs
        if (!(auditOnly || loadOneDrive) && !appProps.hideTabs.search && !isEmpty(appProps.searchProps) && appProps.searchProps.key) {
            items.push({
                tabName: appProps.searchProps.tabName || "Search Property",
                onRender: (el) => {
                    // Render the tab
                    this._tabSearch = new SearchPropTab(el, appProps.searchProps);
                }
            });
        }
        if (!(auditOnly || loadOneDrive) && !appProps.hideTabs.management) {
            items.push({
                tabName: "Management",
                onRender: (el) => {
                    // Render the tab
                    this._tabManagement = new ManagementTab(el, appProps.siteProps, appProps.maxStorageSize, appProps.maxStorageDesc);
                }
            });
        }
        if (!(auditOnly || loadOneDrive) && !appProps.hideTabs.features) {
            items.push({
                tabName: "Features",
                onRender: (el) => {
                    // Render the tab
                    this._tabFeatures = new FeaturesTab(el, appProps.siteProps);
                }
            });
        }
        if (!(auditOnly || loadOneDrive) && !appProps.hideTabs.appPermissions) {
            items.push({
                tabName: "App Permissions",
                onRender: (el) => {
                    // Render the tab
                    this._tabAppPermissions = new AppPermissionsTab(el);
                }
            });
        }
        if (!loadOneDrive && !appProps.hideTabs.lists) {
            items.push({
                tabName: "Lists/Libraries",
                onRender: (el) => {
                    // Render the tab
                    this._tabLists = new ListsTab(el, appProps);
                }
            });
        }
        if (loadOneDrive || !appProps.hideTabs.auditTools) {
            items.push({
                isActive: loadOneDrive,
                tabName: "Audit Tools",
                onRender: (el) => {
                    // Render the tab
                    this._tabReports = new ReportsTab(el, appProps, loadOneDrive);
                }
            });
        }
        if (!(auditOnly || loadOneDrive) && !appProps.hideTabs.webs) {
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
        }
        if (!(auditOnly || loadOneDrive)) {
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
        }

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
        // Ensure the lists or web tab exists
        if (this._tabLists || this._tabWeb) {
            // Show a loading dialog
            LoadingDialog.setHeader("Loading Web");
            LoadingDialog.setBody("Loading the selected web...");
            LoadingDialog.show();

            // Append the sub-webs
            if (this._tabWeb) {
                this._webRequests = this._webRequests.concat(this._tabWeb.getRequests());
            }

            // Load the web information
            DataSource.loadWebInfo(url).then(() => {
                // See if the tab exists
                if (this._tabWeb) {
                    // Update the tab name
                    this._elWebTab.innerHTML = DataSource.Site.RootWeb.Id == DataSource.Web.Id ? "Top Site" : "Sub Site";

                    // Refresh the tabs
                    this._tabWeb ? this._tabWeb.refresh() : null;
                }

                // Refresh the lists tabs
                this._tabLists ? this._tabLists.loadLists(false) : null;

                // Hide the loading dialog
                LoadingDialog.hide();
            });
        }
    }
}