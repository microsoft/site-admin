import { List } from "dattatable";
import { Components, ContextInfo, Helper, Site, Types, Web, Search, SPTypes } from "gd-sprest-bs";
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
    value: boolean | string;
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
 * Site Information
 */
export interface ISiteInfo {
    Path: string;
    SiteId: string;
    Title: string;
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

    // Gets all web urls for a site collection
    static getAllWebs(url: string): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve) => {
            // Get all the sub sites
            Web(url, { requestDigest: this.SiteContext.FormDigestValue }).query({
                Expand: ["Webs"],
                Select: ["Webs/Id", "Webs/ServerRelativeUrl"]
            }).execute(resp => {
                // Parse the webs
                Helper.Executor(resp.Webs.results, web => {
                    // Append the item
                    this._siteItems.push({
                        text: web.ServerRelativeUrl,
                        value: web.Id
                    });

                    // Return a promise
                    return new Promise(resolve => {
                        // Get the sub sites
                        this.getAllWebs(web.ServerRelativeUrl).then(resolve, resolve);
                    });
                }).then(resolve);
            });
        });
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

    // Loads the site collections for the current user
    private static _sites: ISiteInfo[] = null;
    static get MySites(): ISiteInfo[] { return this._sites; }
    private static _mySiteItems: Components.IDropdownItem[] = null;
    static get MySiteItems(): Components.IDropdownItem[] { return this._mySiteItems; }
    static loadSites(): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // See if the sites have been loaded
            if (this._sites) { resolve(); }

            // Search for the sites
            Search.postQuery<ISiteInfo>({
                getAllItems: true,
                query: {
                    Querytext: "contentclass=sts_site",
                    SelectProperties: {
                        results: [
                            "Path", "SiteId", "Title"
                        ]
                    }
                }
            }).then(search => {
                // Set the sites
                this._sites = search.results;

                // Clear the items
                this._mySiteItems = [];

                // Parse the items
                for (let i = 0; i < search.results.length; i++) {
                    let item = search.results[i];

                    // Add the item
                    this._mySiteItems.push({
                        text: item.Path.replace(document.location.origin, ""),
                        value: item.Path
                    });
                }

                // Sort the items
                this._mySiteItems = this._mySiteItems.sort((a, b) => {
                    if (a.text < b.text) { return -1; }
                    if (a.text > b.text) { return 1; }
                    return 0;
                });

                // Resolve the request
                resolve();
            }, reject);
        });
    }

    // Loads the site collection information
    private static _site: Types.SP.SiteOData = null;
    static get Site(): Types.SP.SiteOData { return this._site; }
    private static _siteContext: Types.SP.ContextWebInformation = null;
    static get SiteContext(): Types.SP.ContextWebInformation { return this._siteContext; }
    static get SiteCustomScriptsEnabled(): boolean { return Helper.hasPermissions(DataSource.Site.RootWeb.EffectiveBasePermissions, SPTypes.BasePermissionTypes.AddAndCustomizePages); }
    private static _siteItems: Components.IDropdownItem[] = null;
    static get SiteItems(): Components.IDropdownItem[] { return this._siteItems; }
    private static loadSiteInfo(): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Load the web
            Site(this.SiteContext.SiteFullUrl, { requestDigest: this.SiteContext.FormDigestValue }).query({
                Expand: ["Features", "RootWeb/AllProperties", "RootWeb/EffectiveBasePermissions", "Usage"],
                Select: [
                    "CommentsOnSitePagesDisabled",
                    "DisableCompanyWideSharingLinks",
                    "HubSiteId",
                    "Id",
                    "IsHubSite",
                    "MediaTranscriptionDisabled",
                    "Owner",
                    "ReadOnly",
                    "RootWeb/Created",
                    "RootWeb/Id",
                    "RootWeb/Title",
                    "RootWeb/WebTemplate",
                    "SandboxedCodeActivationCapability",
                    "SecondaryContact",
                    "ServerRelativeUrl",
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

                // Clear the items
                this._siteItems = [{
                    text: this._site.ServerRelativeUrl,
                    value: this._site.RootWeb.Id
                }];

                // Get all of the sites for this collection
                this.getAllWebs(site.Url).then(() => {
                    // Sort the items
                    this._siteItems = this._siteItems.sort((a, b) => {
                        if (a.text < b.text) { return -1; }
                        if (a.text > b.text) { return 1; }
                        return 0;
                    });

                    // Resolve the request
                    resolve();
                }, reject);
            }, reject);
        });
    }

    // Loads the web information
    private static _web: Types.SP.WebOData = null;
    static get Web(): Types.SP.WebOData { return this._web; }
    static loadWebInfo(url: string): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Load the web
            Web(url, { requestDigest: this.SiteContext.FormDigestValue }).query({
                Expand: ["AllProperties"],
                Select: [
                    "Configuration",
                    "CommentsOnSitePagesDisabled",
                    "Created",
                    "ExcludeFromOfflineClient",
                    "Id",
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

    // Loads the web template
    private static _webTemplates: { [key: string]: string } = {};
    static getWebTemplate(key: string): PromiseLike<string> {
        // Return a promise
        return new Promise(resolve => {
            // See if we already have the web template info
            if (this._webTemplates[key]) { resolve(this._webTemplates[key]); return; }

            // Get the web template
            this.Web.getAvailableWebTemplates(1033).getByName(key).execute(template => {
                // Set the value
                this._webTemplates[key] = template.Title;

                // Resolve the request
                resolve(template.Title);
            }, () => {
                // Resolve the request
                resolve(key);
            });
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
    static hasAppCatalog(): boolean {
        // Parse the site features
        for (let i = 0; i < this.Site.Features.results.length; i++) {
            let feature = this.Site.Features.results[i];

            // See if the site collection app catalog feature is enabled
            if (feature.DefinitionId == "88dc6e04-3256-401b-9851-8e07674bb0d6") { return true; }
        }

        // Not found
        return false;
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

    // Updates the search property
    static updateSearchProp(key: string, value: string): PromiseLike<void> {
        // Return a promise
        return new Promise(resolve => {
            // Ensure the key and value exist
            if (key && value != null) {
                // Update the property
                Helper.setWebProperty(key, value, true, DataSource.Web.Url).then(() => {
                    // Resolve the request
                    resolve();
                }, resolve);
            } else {
                // Resolve the request
                resolve();
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
                    this._siteContext = context.GetContextWebInformation;

                    // Check to see if the user is a SCA
                    Web(this.SiteContext.SiteFullUrl).SiteUsers().getByEmail(ContextInfo.userEmail).execute(user => {
                        // Ensure this is an SCA
                        if (user.IsSiteAdmin) {
                            // Load the web information
                            this.loadWebInfo(this.SiteContext.WebFullUrl).then(() => {
                                // Load the site information
                                this.loadSiteInfo().then(resolve, reject);
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