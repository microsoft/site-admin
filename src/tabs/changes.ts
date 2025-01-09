import { Dashboard } from "dattatable";
import { IRequest } from "../ds";

export interface IChangeRequest {
    property?: string;
    newValue: string | boolean | number;
    oldValue: string | boolean | number;
    request?: IRequest
    scope: "Site" | "Web"
}

/**
 * Changes Tab
 * Displays the changes and confirmation button to apply them.
 */
export class ChangesTab {
    private _el: HTMLElement = null;

    // Constructor
    constructor(el: HTMLElement) {
        this._el = el;
    }

    // Process the requests
    processRequests(requests: IChangeRequest[]) {
        // TODO
    }

    // Set the site changes
    setChanges(requests: IChangeRequest[]) {
        // Clear the element
        while (this._el.firstChild) { this._el.removeChild(this._el.firstChild); }

        // Render a table
        new Dashboard({
            el: this._el,
            navigation: {
                showFilter: false,
                items: [
                    {
                        className: "btn-outline-light",
                        isButton: true,
                        isDisabled: requests.length == 0,
                        text: "Apply Changes",
                        onClick: () => {
                            // Process the requests
                            this.processRequests(requests);
                        }
                    }
                ]
            },
            table: {
                rows: requests,
                columns: [
                    {
                        name: "scope",
                        title: "Scope"
                    },
                    {
                        name: "property",
                        title: "Property",
                        onRenderCell: (el, col, change: IChangeRequest) => {
                            // See if this is a request
                            if (change?.request) {
                                el.innerHTML = change.request.key;
                            }
                        }
                    },
                    {
                        name: "oldValue",
                        title: "Old Value"
                    },
                    {
                        name: "newValue",
                        title: "New Value"
                    },
                    {
                        name: "",
                        title: "Time to Complete",
                        onRenderCell: (el, col, change: IChangeRequest) => {
                            // Set the time to complete
                            if (change?.request) {
                                el.innerHTML = change.request.message;
                            } else {
                                el.innerHTML = "This will be completed immediately.";
                            }
                        }
                    },
                ]
            }
        });
    }
}