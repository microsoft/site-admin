import { List, LoadingDialog } from "dattatable";
import { Components, ContextInfo, GroupSiteManager, Helper, Site, Types, Web, SPTypes, v2, DirectorySession } from "gd-sprest-bs";
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
 * Sensitivity Label
 */
export interface ISensitivityLabel {
    desc: string;
    id: string;
    name: string;
}

/**
 * Request Types
 */
export enum RequestTypes {
    AppCatalog = "App Catalog",
    ClientSideAssets = "Client Side Assets",
    CustomScript = "Custom Script",
    DisableCompanyWideSharingLinks = "Company Wide Sharing Links",
    IncreaseStorage = "Increase Storage",
    LockState = "Lock State",
    NoCrawl = "No Crawl",
    SearchProperty = "Custom Search Property",
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

    // Checks to see if a site is set to read-only and determines if the user is a site admin
    private static checkReadOnlySite(siteUrl: string): PromiseLike<boolean> {
        // Return a promise
        return new Promise(resolve => {
            // Get the site
            Site(siteUrl).query({ Select: ["ReadOnly"] }).execute(site => {
                // See if it's not read only
                if (!site.ReadOnly) {
                    // Resolve the request
                    resolve(false);
                    return;
                }

                // Get the site admins for this site
                Web(siteUrl).SiteUsers().query({
                    Filter: "IsSiteAdmin eq true",
                    OrderBy: ["PrincipalType desc"]
                }).execute(items => {
                    let groupIds = [];

                    // Parse the items
                    for (let i = 0; i < items.results.length; i++) {
                        let item = items.results[i];

                        // See if this is a user
                        if (item.PrincipalType == SPTypes.PrincipalTypes.User) {
                            // See if this is the user
                            if (item.Email == ContextInfo.userEmail) {
                                // Resolve the request
                                resolve(true);
                                return;
                            }
                        }
                        // Else, see if this is a M365 group
                        else if (item.PrincipalType == SPTypes.PrincipalTypes.SecurityGroup) {
                            // Add the group id
                            groupIds.push(this.getGroupId(item.LoginName));
                        }
                    }

                    // See if no groups exist
                    if (groupIds.length == 0) {
                        // Resolve the request
                        resolve(false);
                        return;
                    }

                    // Create the request for getting the M365 group information in a batch request
                    let ds = DirectorySession();

                    // Parse the group ids to see if the user is a site admin
                    let isSiteAdmin = false;
                    groupIds.forEach(groupId => {
                        // See if the user is an member/owner of this group
                        ds.group(groupId).query({
                            Expand: ["members", "owners"],
                            Select: [
                                "id",
                                "members/principalName", "members/id", "members/displayName", "members/mail",
                                "owners/principalName", "owners/id", "owners/displayName", "owners/mail"
                            ]
                        }).batch(group => {
                            // See if we already set the flag
                            if (isSiteAdmin) { return; }

                            // See if the user is a member
                            for (let i = 0; i < group.members.results.length; i++) {
                                if (group.members.results[i].mail == ContextInfo.userEmail) {
                                    // Set the flag
                                    isSiteAdmin = true;
                                    return;
                                }
                            }

                            // See if the user is a owner
                            for (let i = 0; i < group.owners.results.length; i++) {
                                if (group.owners.results[i].mail == ContextInfo.userEmail) {
                                    // Set the flag
                                    isSiteAdmin = true;
                                    return;
                                }
                            }
                        });
                    });

                    // Execute the request
                    ds.execute(() => {
                        // Resolve the request
                        resolve(isSiteAdmin);
                    });
                }, () => {
                    // Resolve the request
                    resolve(false);
                });
            }, () => {
                // Resolve the request
                resolve(false);
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
                onInitialized: resolve
            });
        });
    }

    // Returns the group id from the login name
    static getGroupId(loginName: string): string {
        // Get the group id from the login name
        let userInfo = loginName.split('|');
        let groupInfo = userInfo[userInfo.length - 1];
        let groupId = groupInfo.split('_')[0];

        // Ensure it's a guid and return null if it's not
        return /^[{]?[0-9a-fA-F]{8}(-[0-9a-fA-F]{4}){3}-[0-9a-fA-F]{12}[}]?$/.test(groupId) ? groupId : null;
    }

    // Loads the files for a drive
    static loadFiles(webId: string, listName?: string, folder?: Types.SP.Folder): PromiseLike<Types.Microsoft.Graph.driveItem[]> {
        let files = [];

        // Loads the files for a drive
        let getFiles = (driveId: string, folder: string = "") => {
            let drive = v2.sites({ siteId: this.Site.Id, webId }).drives(driveId);

            // Return a promise
            return new Promise(resolve => {
                // Get the files for the folder
                let driveFolder = folder ? drive.getFolder(folder) : drive.root();
                driveFolder.children().query({
                    GetAllItems: true,
                    Select: ["driveId", "file", "folder", "id", "name", "parentReference", "sensitivityLabel", "webUrl"],
                    Top: 5000
                }).execute(resp => {
                    // Parse the items
                    Helper.Executor(resp.results, driveItem => {
                        // See if this is a file
                        if (driveItem.file) {
                            // Append the file
                            files.push(driveItem);
                        } else {
                            // Get the items for this folder
                            return getFiles(driveId, (folder ? folder + "/" : "") + driveItem.name);
                        }
                    }).then(() => {
                        // Resolve the request
                        resolve(files);
                    });
                });
            });
        }

        // Return a promise
        return new Promise((resolve, reject) => {
            // Get the libraries for this site
            v2.sites({ siteId: DataSource.Site.Id, webId }).drives().execute(resp => {
                let drives: Types.Microsoft.Graph.drive[] = [];

                // See if we are searching for a specific library
                if (listName) {
                    // Find the target drive
                    let drive = resp.results.find(a => { return a.name == listName; });
                    if (drive) { drives.push(drive); }
                } else {
                    // Parse the drives
                    for (let i = 0; i < resp.results.length; i++) {
                        let drive = resp.results.find(a => { return a.name == drives[i].name; });
                        if (drive) { drives.push(drive); }
                    }
                }

                // Parse the drives
                Helper.Executor(drives, drive => {
                    return getFiles(drive.id, folder?.Name);
                }).then(() => { resolve(files); });
            }, reject);
        });
    }

    // Loads the sensitivity labels for the current user
    private static _sensitivityLabels: ISensitivityLabel[] = null;
    static get HasSensitivityLabels(): boolean { return this._sensitivityLabels?.length > 0; }
    static get SensitivityLabels(): ISensitivityLabel[] { return this._sensitivityLabels; }
    private static _sensitivityLabelItems: Components.IDropdownItem[] = null;
    static get SensitivityLabelItems(): Components.IDropdownItem[] { return this._sensitivityLabelItems; }
    static getSensitivityLabel(labelId: string): string {
        // Find the item
        let item = this.SensitivityLabels.filter(a => { return a.id == labelId; })[0];

        // Return the value
        return item ? item.name : labelId;
    }
    private static loadSensitivityLabels() {
        // Return a promise
        return new Promise(resolve => {
            // Load the group context
            GroupSiteManager().getGroupCreationContext().execute(resp => {
                // Clear the labels
                this._sensitivityLabels = [];
                this._sensitivityLabelItems = [
                    {
                        text: "",
                        value: null
                    }
                ];

                // Parse the sensitivity labels
                for (let i = 0; i < resp.DataClassificationOptionsNew.results.length; i++) {
                    let result = resp.DataClassificationOptionsNew.results[i];
                    let desc = resp.ClassificationDescriptionsNew.results.filter(i => { return i.Key == result.Value; })[0];

                    // Append the label and item
                    this._sensitivityLabels.push({
                        desc: desc ? desc.Value : "",
                        id: result.Key,
                        name: result.Value
                    });
                    this._sensitivityLabelItems.push({
                        text: result.Value,
                        value: result.Key
                    });
                }

                // Resolve the request
                resolve(null);
            }, resolve);
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
                    "SensitivityLabelId",
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

                // Load additional information
                Promise.all([
                    // Load the client side assets
                    this.loadClientSideAssets(),
                    // Get all of the sites for this collection
                    this.getAllWebs(site.Url)
                ]).then(() => {
                    // Sort the items
                    this._siteItems = this._siteItems.sort((a, b) => {
                        if (a.text < b.text) { return -1; }
                        if (a.text > b.text) { return 1; }
                        return 0;
                    });

                    // Resolve the request
                    resolve();
                }).catch(reject);
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
                    "NoCrawl",
                    "SearchScope",
                    "SensitivityLabelId",
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

    // Gets the url from the query string
    static getUrlFromQS() {
        // Get the id from the querystring
        let qs = document.location.search.split('?');
        qs = qs.length > 1 ? qs[1].split('&') : [];
        for (let i = 0; i < qs.length; i++) {
            let qsItem = qs[i].split('=');
            let key = qsItem[0].toLowerCase();
            let value = qsItem[1];

            // See if this is the "site" key
            if (key == "site") {
                // Return the url
                return value;
            }
        }
    }

    // Determines if the app catalog has bypassed the block download policy
    private static _hasBypassBlockDownloadPolicy: boolean;
    static get HasBypassBlockDownloadPolicy(): boolean { return this._hasBypassBlockDownloadPolicy; }
    private static loadClientSideAssets(): PromiseLike<any> {
        // Clear the flag
        this._hasBypassBlockDownloadPolicy = false;

        // Return a promise
        return new Promise((resolve) => {
            // Get the client side assets library
            Web(this.SiteContext.SiteFullUrl).Lists("Client Side Assets").query({ Select: ["ExemptFromBlockDownloadOfNonViewableFiles"] }).execute(list => {
                // Set the flag
                this._hasBypassBlockDownloadPolicy = list.ExemptFromBlockDownloadOfNonViewableFiles;

                // Resolve the request
                resolve(null);
            }, resolve);
        });
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
    static init(): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Init the security
            Security.init().then(() => {
                Promise.all([
                    // Load the list
                    this.loadList(),
                    // Load the sensitivity labels
                    this.loadSensitivityLabels()
                ]).then(() => {
                    // See if a url exists in the query string
                    let url = this.getUrlFromQS();
                    if (url) {
                        // Update the loading dialog
                        LoadingDialog.setHeader("Loading Site");
                        LoadingDialog.setBody("Validating the site '" + url + "'");

                        // Validate the url and resolve the request
                        this.validate(url).then(resolve, resolve);
                    } else {
                        // Resolve the request
                        resolve();
                    }
                }, reject);
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

    // Search Property
    private static _searchPropItems: Components.IDropdownItem[] = null;
    static get SearchPropItems(): Components.IDropdownItem[] { return this._searchPropItems; }
    static set SearchPropItems(strCSV: string) {
        // Clear the items
        this._searchPropItems = [{ text: "", value: null }];

        // Parse the values
        let values = strCSV.split(',');
        for (let i = 0; i < values.length; i++) {
            let value = values[i].trim();
            if (value) {
                // add the item
                this._searchPropItems.push({
                    text: value,
                    value
                });
            }
        }
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
                    // Set the site context
                    this._siteContext = context.GetContextWebInformation;

                    // Get the sharing settings for the web
                    Web.getSharingSettings({
                        objectUrl: this.SiteContext.SiteFullUrl
                    }).execute(settings => {
                        // See if this is the site admin
                        if (settings.IsUserSiteAdmin) {
                            // Load the web information
                            this.loadWebInfo(this.SiteContext.WebFullUrl).then(() => {
                                // Load the site information
                                this.loadSiteInfo().then(resolve, reject);
                            }, reject);
                        } else {
                            // Check to see if this is a read-only site and determine if the user is an admin
                            this.checkReadOnlySite(this.SiteContext.SiteFullUrl).then(isAdmin => {
                                if (isAdmin) {
                                    // Load the web information
                                    this.loadWebInfo(this.SiteContext.WebFullUrl).then(() => {
                                        // Load the site information
                                        this.loadSiteInfo().then(resolve, reject);
                                    }, reject);
                                } else {
                                    // Reject the request
                                    reject("Site exists, but you are not the administrator. Please have the site administrator submit the request.");
                                }
                            });
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