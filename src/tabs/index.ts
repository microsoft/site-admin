import { Components } from "gd-sprest-bs";
import { AppPermissionsTab } from "./appPermissions";
import { ChangesTab } from "./changes";
import { FeaturesTab } from "./features";
import { InfoTab } from "./info";
import { ManagementTab } from "./management";
import { WebsTab } from "./webs";
import { ITab } from "./tab.d";

/**
 * Tabs
 */
export class Tabs {
    private _el: HTMLElement = null;
    private _changesTab: ChangesTab = null;
    private _tabs: ITab[] = [];

    // Constructor
    constructor(el: HTMLElement) {
        this._el = el;

        // Render the tabs
        this.render();
    }

    // Renders the tabs
    private render() {
        // Render the tabs
        Components.ListGroup({
            el: this._el,
            isHorizontal: true,
            isTabs: true,
            colWidth: 12,
            items: [
                {
                    tabName: "Information",
                    isActive: true,
                    onRender: (el) => {
                        // Render the tab
                        this._tabs.push(new InfoTab(el));
                    }
                },
                {
                    tabName: "Management",
                    onRender: (el) => {
                        // Render the tab
                        this._tabs.push(new ManagementTab(el));
                    }
                },
                {
                    tabName: "Features",
                    onRender: (el) => {
                        // Render the tab
                        this._tabs.push(new FeaturesTab(el));
                    }
                },
                {
                    tabName: "App Permissions",
                    onRender: (el) => {
                        // Render the tab
                        this._tabs.push(new AppPermissionsTab(el));
                    }
                },
                {
                    tabName: "Web(s)",
                    onRender: (el) => {
                        // Render the tab
                        this._tabs.push(new WebsTab(el));
                    }
                },
                {
                    tabName: "Changes",
                    onRender: (el) => {
                        // Render the changes
                        this._changesTab = new ChangesTab(el);
                    }
                }
            ]
        });
    }
}