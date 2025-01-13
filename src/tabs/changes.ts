import { Dashboard, Modal, LoadingDialog } from "dattatable";
import { Components, Helper, Web } from "gd-sprest-bs";
import { DataSource, IRequest } from "../ds";
import { isEmpty } from "./common";

export interface IChangeRequest {
    property?: string;
    newValue: string | boolean | number;
    oldValue: string | boolean | number;
    request?: IRequest
    response?: string;
    scope: "Site" | "Web",
    url: string;
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
    processRequests(changes: IChangeRequest[]) {
        let requests: IRequest[] = [];
        let siteProps = {};

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
                }
            }
        }

        // Set the status
        LoadingDialog.setBody("Applying the site changes...");

        // Update the site properties
        this.updateSite(siteProps).then((siteUpdateError) => {
            // Set the status
            LoadingDialog.setBody("Applying the web changes...");

            // Update the web properties
            this.updateWeb(changes).then(() => {
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
                        } else if (change.scope == "Site") {
                            // Set the response
                            change.response = `The request ${siteUpdateError ? "failed to be updated" : "was completed successfully"}.`
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

                            // Show a confirmation dialog
                            Modal.clear();
                            Modal.setHeader("Update Completed");
                            Modal.setBody("The changes have completed. Please refer to the table for additional information of the requests.");
                            Components.Button({
                                el: Modal.FooterElement,
                                text: "Close",
                                onClick: () => { Modal.hide(); }
                            });
                            Modal.show();
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
                        name: "url",
                        title: "Url"
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
                        name: "",
                        title: "Value",
                        onRenderCell: (el, col, change: IChangeRequest) => {
                            // Render the old/new values
                            el.innerHTML = `
                                <b>Current Value:</b> ${change.oldValue}
                                <br/>
                                <b>New Value:</b> ${change.newValue}
                            `;
                        }
                    },
                    {
                        name: "",
                        title: "Status",
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

    // Update the site properties
    private updateSite(props): PromiseLike<boolean> {
        // Return a promise
        return new Promise(resolve => {
            // See if the properties are empty
            if (isEmpty(props)) {
                // Resolve the request
                resolve(false);
                return;
            }

            // Save the changes
            DataSource.Site.update(props).execute(
                () => { resolve(false) },
                () => { resolve(true) }
            );

        });
    }

    // Update the web properties
    private updateWeb(requests: IChangeRequest[]) {
        // Return a promise
        return new Promise(resolve => {
            // Parse the requests
            Helper.Executor(requests, request => {
                // Return a promise
                return new Promise(resolve => {
                    let props = {};
                    props[request.property] = request.newValue;

                    // Save the changes
                    Web(request.url, { requestDigest: DataSource.SiteContext.FormDigestValue }).update(props).execute(
                        () => {
                            // Update the request
                            request.response = "The request was completed successfully.";
                            resolve(null);
                        },
                        () => {
                            // Update the request
                            request.response = "The request failed to be updated.";
                            resolve(null);
                        }
                    );
                });
            }).then(resolve);
        });
    }
}