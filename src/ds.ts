import { List } from "dattatable";
import { Components, ContextInfo, Site, Types, Web } from "gd-sprest-bs";
import { Security } from "./security";
import Strings from "./strings";

/**
 * List Item
 * Add your custom fields here
 */
export interface IListItem extends Types.SP.ListItem {
    AuthorId: number;
    Author: { Id: number; Title: string; }
    Status: string;
}

/**
 * Site Information
 */
export interface ISiteInfo {
    site: Types.SP.SiteOData;
    web: Types.SP.WebOData;
}

/**
 * Data Source
 */
export class DataSource {
    // Azure Function Url
    private static _azureFunctionUrl: string = null;
    static get AzureFunctionEnabled(): boolean { return this._azureFunctionUrl ? true : false; }
    static processRequest(itemId: number): PromiseLike<string> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Create the request
            let xhr = new XMLHttpRequest();
            xhr.open("POST", this._azureFunctionUrl, true);

            // Set the header
            xhr.setRequestHeader("Content-Type", "application/json");

            // Set the event
            xhr.onreadystatechange = (ev) => {
                if (xhr.readyState !== 4) { return; }

                // See if it was successful
                if (xhr.status === 200) {
                    // Resolve the request                
                    resolve(xhr.responseText);
                } else {
                    // Reject the request
                    reject(xhr.responseText);
                }
            }

            // Send the request
            xhr.send(JSON.stringify({ requestId: itemId }));
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

    // Loads the site collection information
    private static loadSiteInfo(context: Types.SP.ContextWebInformation): PromiseLike<Types.SP.SiteOData> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Load the web
            Site(context.SiteFullUrl, { requestDigest: context.FormDigestValue }).query({
                Select: [
                    "CommentsOnSitePagesDisabled",
                    "DisableCompanyWideSharingLinks",
                    "MediaTranscriptionDisabled",
                    "Owner",
                    "SandboxedCodeActivationCapability",
                    "SecondaryContact",
                    "ShareByEmailEnabled",
                    "ShowPeoplePickerSuggestionsForGuestUsers",
                    "SocialBarOnSitePagesDisabled",
                    "StatusBarLink",
                    "StatusBarText"
                ]
            }).execute(resolve, reject);
        });
    }

    // Loads the web information
    private static loadWebInfo(context: Types.SP.ContextWebInformation): PromiseLike<Types.SP.WebOData> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Load the web
            Web(context.WebFullUrl, { requestDigest: context.FormDigestValue }).query({
                Expand: ["AllProperties"],
                Select: [
                    "Configuration",
                    "CommentsOnSitePagesDisabled",
                    "ExcludeFromOfflineClient",
                    "SearchScope",
                    "Url",
                    "WebTemplate"
                ]
            }).execute(resolve, reject);
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

    // Initializes the application
    static init(azureFunctionUrl: string): PromiseLike<any> {
        // Set the url
        this._azureFunctionUrl = azureFunctionUrl;

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
    static validate(webUrl: string): PromiseLike<ISiteInfo> {
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
                            this.loadWebInfo(context.GetContextWebInformation).then(web => {
                                // Load the site information
                                this.loadSiteInfo(context.GetContextWebInformation).then(site => {
                                    // Resolve the request
                                    resolve({ site, web });
                                }, reject);
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