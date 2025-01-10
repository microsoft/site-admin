import { Dashboard, LoadingDialog } from "dattatable";
import { DataSource, IRequest } from "../ds";

export interface IChangeRequest {
    property?: string;
    newValue: string | boolean | number;
    oldValue: string | boolean | number;
    request?: IRequest
    response?: string;
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

    // Sees if the properties is empty
    isEmpty(obj) {
        // Return true if null
        if (obj == null) { return true; }

        // Parse the properties
        for (var prop in obj) {
            // See if the object has properties
            if (Object.prototype.hasOwnProperty.call(obj, prop)) { return false; }
        }

        // Object is empty
        return true;
    }

    // Process the requests
    processRequests(changes: IChangeRequest[]) {
        let requests: IRequest[] = [];
        let siteProps = {};
        let webProps = {};

        // Parse the requests
        for (let i = 0; i < changes.length; i++) {
            let change = changes[i];

            // See if this is a request
            if (change.request) {
                // Add the request
                requests.push(change.request);
            } else {
                // Set the site/web property
                if (change.scope == "Site") {
                    siteProps[change.property] = change.newValue;
                } else {
                    webProps[change.property] = change.newValue;
                }
            }
        }

        // Set the status
        LoadingDialog.setBody("Applying the site changes...");

        // Update the site properties
        this.updateProperties(siteProps, "Site").then((siteUpdateError) => {
            // Set the status
            LoadingDialog.setBody("Applying the web changes...");

            // Update the web properties
            this.updateProperties(webProps, "Web").then((webUpdateError) => {
                // Set the status
                LoadingDialog.setBody("Creating the request items...");

                // Add the requests
                DataSource.addRequest(DataSource.Site.Url, requests).then((responses) => {
                    // Parse the changes
                    for (let i = 0; i < changes.length; i++) {
                        let change = changes[i];

                        // See if this is the request
                        if (change.request) {
                            // Parse the response
                            for (let i = 0; i < responses.length; i++) {
                                let response = responses[i];

                                // See if this is the response
                                if (change.request.key == response.key) {
                                    // Set the response
                                    change.response = response.message;
                                    break;
                                }
                            }
                        } else {
                            let errorFl = change.scope == "Site" ? siteUpdateError : webUpdateError;

                            // Set the response
                            change.response = `The request ${errorFl ? "failed to be updated" : "was completed successfully"}.`
                        }
                    }

                    // Set the changes
                    this.setChanges(changes, true);

                    // Hide the loading dialog
                    LoadingDialog.hide();
                });
            });
        });
    }

    // Set the site changes
    setChanges(requests: IChangeRequest[], isResponse: boolean = false) {
        // Clear the element
        while (this._el.firstChild) { this._el.removeChild(this._el.firstChild); }

        // Render a table
        new Dashboard({
            el: this._el,
            navigation: {
                showFilter: false,
                showSearch: false,
                itemsEnd: [
                    {
                        className: "btn-outline-light",
                        isButton: true,
                        isDisabled: requests.length == 0,
                        text: "Apply Changes",
                        onClick: () => {
                            // Show a loading dialog
                            LoadingDialog.setHeader("Applying Changes");
                            LoadingDialog.show();

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
                        title: isResponse ? "Status" : "Time to Complete",
                        onRenderCell: (el, col, change: IChangeRequest) => {
                            // See if this is a response
                            if (isResponse) {
                                // Set the response
                                el.innerHTML = change.response;
                            } else {
                                // Set the time to complete
                                if (change?.request) {
                                    el.innerHTML = change.request.message;
                                } else {
                                    el.innerHTML = "This will be completed immediately.";
                                }

                            }
                        }
                    },
                ]
            }
        });
    }

    // Update the properties
    private updateProperties(props, scope: "Site" | "Web"): PromiseLike<boolean> {
        // Return a promise
        return new Promise(resolve => {
            // See if the properties are empty
            if (this.isEmpty(props)) {
                // Resolve the request
                resolve(false);
                return;
            }

            // See if we are updating the site
            if (scope == "Site") {
                // Save the changes
                DataSource.Site.update(props).execute(
                    () => { resolve(false) },
                    () => { resolve(true) }
                );
            } else {
                // Save the changes
                DataSource.Web.update(props).execute(
                    () => { resolve(false) },
                    () => { resolve(true) }
                );
            }
        });
    }
}