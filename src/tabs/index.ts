import { Components } from "gd-sprest-bs";
import { AppPermissionsTab } from "./appPermissions";
import { ChangesTab } from "./changes";
import { FeaturesTab } from "./features";
import { InfoTab } from "./info";
import { ManagementTab } from "./management";
import { SearchPropTab, ISearchProp } from "./searchProp";
import { WebsTab } from "./webs";

/**
 * Tabs
 */
export class Tabs {
    private _el: HTMLElement = null;
    private _tabAppPermissions: AppPermissionsTab = null;
    private _tabChanges: ChangesTab = null;
    private _tabFeatures: FeaturesTab = null;
    private _tabManagement: ManagementTab = null;
    private _tabSearch: SearchPropTab = null;
    private _tabWebs: WebsTab = null;

    // Constructor
    constructor(el: HTMLElement, disableSiteProps: string[] = [], disableWebProps: string[] = [], searchProp: ISearchProp) {
        this._el = el;

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
        if (searchProp) {
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
                this._tabWebs = new WebsTab(el, disableWebProps);
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
                    this._tabWebs.getRequests()
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
}