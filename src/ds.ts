import { List } from "dattatable";
import { Components, ContextInfo, Helper, Site, SPTypes, Types, Web, v2 } from "gd-sprest-bs";
import { Security } from "./security";
import Strings from "./strings";

/**
 * List Item
 * Add your custom fields here
 */
export interface IListItem extends Types.SP.ListItem {
    AuthorId: number;
    Author: { Id: number; Title: string; }
    ProcessFlag: boolean;
    RequestType: string;
    RequestValue: string;
    Status: string;
}

/**
 * Request for creating an item
 */
export interface IRequest {
    key: string;
    message: string;
    value: boolean;
}

/**
 * Response from creating an item
 */
export interface IResponse {
    errorFl: boolean;
    key: string;
    message: string;
    value: any;
}

/**
 * Request Types
 */
export enum RequestTypes {
    AppCatalog = "App Catalog",
    CustomScript = "Custom Script",
    DisableCompanyWideSharingLinks = "Company Wide Sharing Links",
    IncreaseStorage = "Increase Storage",
    LockState = "Lock State",
    TeamsConnected = "Teams Connected"
}

/**
 * Data Source
 */
export class DataSource {
    // Method to process requests to add to the list
    static addRequest(url: string, requests: IRequest[]): PromiseLike<IResponse[]> {
        let responses: IResponse[] = [];

        // See if any requests exist
        if (requests.length == 0) {
            // Resolve the request
            return Promise.resolve(responses);
        }

        // Return a promise
        return new Promise(resolve => {
            // Parse the requests
            Helper.Executor(requests, request => {
                // Return a promise
                return new Promise(resolve => {
                    // Create the item
                    this.List.createItem({
                        ProcessFlag: true,
                        RequestType: request.key,
                        RequestValue: request.value + "",
                        Title: url
                    }).then(
                        // Success
                        (item) => {
                            // Add the response
                            responses.push({
                                errorFl: false,
                                key: request.key,
                                message: request.message,
                                value: item.RequestType
                            });

                            // Try the next request
                            resolve(null);
                        },

                        // Error
                        ex => {
                            // Add the response
                            responses.push({
                                errorFl: true,
                                key: request.key,
                                message: "There was an error creating the request. Please refresh and try again.",
                                value: request
                            });

                            // Try the next request
                            resolve(null);
                        }
                    );
                });
            }).then(() => {
                // Resolve the request
                resolve(responses);
            });
        });
    }

    // Formats a value to bytes
    static formatBytes(bytes, decimals = 2) {
        if (bytes === 0) return '0 Bytes';
        const k = 1024;
        const dm = decimals < 0 ? 0 : decimals;
        const sizes = ['Bytes', 'KB', 'MB', 'GB', 'TB', 'PB', 'EB', 'ZB', 'YB'];
        const i = Math.floor(Math.log(bytes) / Math.log(k));
        return parseFloat((bytes / Math.pow(k, i)).toFixed(dm)) + ' ' + sizes[i];
    }

    // List Items
    static get ListItems(): IListItem[] { return this.List.Items; }

    // List
    private static _list: List<IListItem> = null;
    static get List(): List<IListItem> { return this._list; }
    private static loadList(): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // See if this is not an admin
            let Filter = null;
            if (!Security.IsAdmin) {
                // Set the filter
                Filter = "AuthorId eq " + ContextInfo.userId;
            }

            // Initialize the list
            this._list = new List<IListItem>({
                listName: Strings.Lists.Main,
                itemQuery: {
                    Filter,
                    Expand: ["Author"],
                    OrderBy: ["Title"],
                    GetAllItems: true,
                    Top: 5000,
                    Select: ["*", "Author/Id", "Author/Title"]
                },
                onInitError: reject,
                onInitialized: () => {
                    // Load the status filters
                    this.loadStatusFilters();

                    // Resolve the request
                    resolve();
                }
            });
        });
    }

    // Loads the site collection information
    private static _site: Types.SP.SiteOData = null;
    static get Site(): Types.SP.SiteOData { return this._site; }
    private static loadSiteInfo(context: Types.SP.ContextWebInformation): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Load the web
            Site(context.SiteFullUrl, { requestDigest: context.FormDigestValue }).query({
                Expand: ["RootWeb/EffectiveBasePermissions", "Usage"],
                Select: [
                    "CommentsOnSitePagesDisabled",
                    "DisableCompanyWideSharingLinks",
                    "GroupId",
                    "MediaTranscriptionDisabled",
                    "Owner",
                    "ReadOnly",
                    "SandboxedCodeActivationCapability",
                    "SecondaryContact",
                    "ShareByEmailEnabled",
                    "ShowPeoplePickerSuggestionsForGuestUsers",
                    "SocialBarOnSitePagesDisabled",
                    "StatusBarLink",
                    "StatusBarText",
                    "Url",
                    "WriteLocked"
                ]
            }).execute(site => {
                // Save the reference and resolve the request
                this._site = site;
                resolve();
            }, reject);
        });
    }

    // Loads the web information
    private static _web: Types.SP.WebOData = null;
    static get Web(): Types.SP.WebOData { return this._web; }
    private static loadWebInfo(context: Types.SP.ContextWebInformation): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Load the web
            Web(context.WebFullUrl, { requestDigest: context.FormDigestValue }).query({
                Expand: ["AllProperties"],
                Select: [
                    "Configuration",
                    "CommentsOnSitePagesDisabled",
                    "Created",
                    "ExcludeFromOfflineClient",
                    "SearchScope",
                    "Title",
                    "Url",
                    "WebTemplate"
                ]
            }).execute(web => {
                // Save the reference and resolve the request
                this._web = web;
                resolve();
            }, reject);
        });
    }

    // Status Filters
    private static _statusFilters: Components.ICheckboxGroupItem[] = null;
    static get StatusFilters(): Components.ICheckboxGroupItem[] { return this._statusFilters; }
    static loadStatusFilters() {
        let items: Components.ICheckboxGroupItem[] = [];

        // Parse the choices
        let fld: Types.SP.FieldChoice = this.List.getField("Status");
        for (let i = 0; i < fld.Choices.results.length; i++) {
            // Add an item
            items.push({
                label: fld.Choices.results[i],
                type: Components.CheckboxGroupTypes.Switch
            });
        }

        // Set the filters and resolve the promise
        this._statusFilters = items;
    }

    // Gets the item id from the query string
    static getItemIdFromQS() {
        // Get the id from the querystring
        let qs = document.location.search.split('?');
        qs = qs.length > 1 ? qs[1].split('&') : [];
        for (let i = 0; i < qs.length; i++) {
            let qsItem = qs[i].split('=');
            let key = qsItem[0];
            let value = qsItem[1];

            // See if this is the "id" key
            if (key == "ID") {
                // Return the item
                return parseInt(value);
            }
        }
    }

    // Determines if the site collection has an app catalog
    static hasAppCatalog(url: string): PromiseLike<boolean> {
        // Return a promise
        return new Promise((resolve) => {
            // Query the root web's lists
            Web(url).Lists().query({ Filter: "BaseTemplate eq " + SPTypes.ListTemplateType.AppCatalog }).execute(
                // Success
                (lists) => {
                    // Exists
                    resolve(lists.results.length > 0);
                },

                () => {
                    // Doesn't exist
                    resolve(false);
                }
            )
        });
    }

    // Initializes the application
    static init(): PromiseLike<any> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Init the security
            Security.init().then(() => {
                // Load the list
                this.loadList().then(resolve, reject);
            }, reject);
        });
    }

    // Refreshes the list data
    static refresh(itemId?: number): PromiseLike<IListItem | IListItem[]> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // See if an item id exists
            if (itemId > 0) {
                // Refresh the item
                DataSource.List.refreshItem(itemId).then(resolve, reject);
            } else {
                // Refresh the data
                DataSource.List.refresh().then(resolve, reject);
            }
        });
    }

    // Validates that the user is an SCA of the site
    static validate(webUrl: string): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Get the web context
            ContextInfo.getWeb(webUrl).execute(
                context => {
                    // Check to see if the user is a SCA
                    Web(context.GetContextWebInformation.SiteFullUrl).SiteUsers().getByEmail(ContextInfo.userEmail).execute(user => {
                        // Ensure this is an SCA
                        if (user.IsSiteAdmin) {
                            // Load the web information
                            this.loadWebInfo(context.GetContextWebInformation).then(() => {
                                // Load the site information
                                this.loadSiteInfo(context.GetContextWebInformation).then(resolve, reject);
                            }, reject);
                        } else {
                            // Not the SCA of the site
                            reject("Site exists, but you are not the administrator. Please have the site administrator submit the request.");
                        }
                    });
                },

                (ex: any) => {
                    // See if it's a permission issue
                    if (ex.status == 403) {
                        // User doesn't have access to site
                        reject("Error the site exists, but you do not have permissions to it.");
                    }
                    // Else, the site doesn't exist
                    else {
                        // Site doesn't exist
                        reject("Error the site doesn't exist. Please check the url and try again.");
                    }
                }
            );
        });
    }
}