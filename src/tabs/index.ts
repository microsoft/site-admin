import { LoadingDialog } from "dattatable";
import { Components } from "gd-sprest-bs";
import { AppPermissionsTab } from "./appPermissions";
import { ChangesTab, IChangeRequest } from "./changes";
import { isEmpty } from "./common";
import { DataSource } from "../ds";
import { FeaturesTab } from "./features";
import { InfoTab } from "./info";
import { ManagementTab } from "./management";
import { SearchPropTab, ISearchProp } from "./searchProp";
import { WebTab } from "./web";

/**
 * Tabs
 */
export class Tabs {
    private _el: HTMLElement = null;
    private _webRequests: IChangeRequest[];
    private _tabAppPermissions: AppPermissionsTab = null;
    private _tabChanges: ChangesTab = null;
    private _tabFeatures: FeaturesTab = null;
    private _tabManagement: ManagementTab = null;
    private _tabSearch: SearchPropTab = null;
    private _tabWeb: WebTab = null;

    // Constructor
    constructor(el: HTMLElement, disableSiteProps: string[] = [], disableWebProps: string[] = [], searchProp: ISearchProp) {
        this._el = el;
        this._webRequests = [];

        // Render the tabs
        this.render(disableSiteProps, disableWebProps, searchProp);
    }

    // Renders the tabs
    private render(disableSiteProps: string[] = [], disableWebProps: string[] = [], searchProp: ISearchProp) {
        // Set the items
        let items: Components.IListGroupItem[] = [
            {
                tabName: "Information",
                isActive: true,
                onRender: (el) => {
                    // Render the tab
                    new InfoTab(el, disableSiteProps);
                }
            },
            {
                tabName: "Management",
                onRender: (el) => {
                    // Render the tab
                    this._tabManagement = new ManagementTab(el, disableSiteProps);
                }
            },
            {
                tabName: "Features",
                onRender: (el) => {
                    // Render the tab
                    this._tabFeatures = new FeaturesTab(el, disableSiteProps);
                }
            },
            {
                tabName: "App Permissions",
                onRender: (el) => {
                    // Render the tab
                    this._tabAppPermissions = new AppPermissionsTab(el);
                }
            }
        ];

        // See if we are customizing a search property
        if (!isEmpty(searchProp)) {
            items.push({
                tabName: "Search Property",
                onRender: (el) => {
                    // Render the tab
                    this._tabSearch = new SearchPropTab(el, searchProp);
                }
            });
        }

        // Add the webs
        items.push({
            tabName: "Web(s)",
            onRender: (el) => {
                // Render the tab
                this._tabWeb = new WebTab(el, disableWebProps);
            }
        });

        // Add the changes
        items.push({
            tabName: "Changes",
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

                // Set the changes
                this._tabChanges.setChanges(changes);
            }
        });

        // Render the tabs
        Components.ListGroup({
            el: this._el,
            isHorizontal: true,
            isTabs: true,
            colWidth: 12,
            items
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
            // Refresh the web tab
            this._tabWeb.refresh();

            // Hide the loading dialog
            LoadingDialog.hide();
        });
    }
}