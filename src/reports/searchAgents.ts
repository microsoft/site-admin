import { Dashboard, Documents, LoadingDialog, Modal } from "dattatable";
import { Components, Helper, SPTypes, Types, Web } from "gd-sprest-bs";
import { DataSource } from "../ds";
import { ExportCSV } from "./exportCSV";

interface IAgentItem {
    FileName?: string;
    FileUrl?: string;
    ItemId?: number;
    ListId?: string;
    ListName?: string;
    ListUrl?: string;
    WebTitle: string;
    WebUrl: string;
}

const CSVFields = [
    "FileName", "ListName", "WebTitle", "ItemId",
    "ListId", "FileUrl", "ListUrl", "WebUrl"
]

export class SearchAgents {
    private static _dashboard: Dashboard = null;
    private static _elSubNav: HTMLElement = null;
    private static _items: IAgentItem[] = null;
    private static _loadOneDrive: boolean = null;
    private static _stopFl: boolean = false;

    // Analyzes a library
    private static analyzeLibrary(web: Types.SP.WebOData, lib: Types.SP.ListOData): PromiseLike<void> {
        // Return a promise
        return new Promise(resolve => {
            let ctrFiles = 0;

            // Set the status
            this._elSubNav.children[1].innerHTML = `Loading Files`;

            // Get the items for this library
            DataSource.loadItems({
                webUrl: web.Url,
                listId: lib.Id,
                isOnedrive: this._loadOneDrive,
                query: {
                    Select: ["Id", "FileLeafRef", "FileRef", "File_x0020_Type"],
                },
                onItem: (item => {
                    // Set the status
                    this._elSubNav.children[1].innerHTML = `Files Loaded: ${++ctrFiles} of ${lib.ItemCount}`;

                    // See if this is an agent
                    if (item["File_x0020_Type"] == "agent") {
                        // Add this item
                        this._items.push({
                            FileName: item["FileRef"],
                            FileUrl: item["FileLeafRef"],
                            ItemId: item.Id,
                            ListId: lib.Id,
                            ListName: lib.Title,
                            ListUrl: lib.RootFolder.ServerRelativeUrl,
                            WebTitle: web.Title,
                            WebUrl: web.Url
                        });
                    }
                })
            }).then(() => { resolve(); });
        });
    }

    // Analyzes a site
    private static analyzeSite(web: Types.SP.WebOData): PromiseLike<void> {
        // Return a promise
        return new Promise(resolve => {
            // Show a dialog
            this._elSubNav.children[1].innerHTML = `Getting Libraries...`;

            // Get the libraries
            let site = this._loadOneDrive ? Web.getOneDrive() : Web(web.Url, { requestDigest: DataSource.SiteContext.FormDigestValue });
            site.Lists().query({
                Filter: `BaseTemplate eq ${SPTypes.ListTemplateType.DocumentLibrary} or BaseTemplate eq ${SPTypes.ListTemplateType.MySiteDocumentLibrary} or BaseTemplate eq ${SPTypes.ListTemplateType.PageLibrary}`,
                Expand: ["RootFolder"],
                Select: ["Id", "Title", "BaseTemplate", "ItemCount", "RootFolder/ServerRelativeUrl"]
            }).execute(resp => {
                let ctrList = 0;
                let siteText = this._elSubNav.children[0].innerHTML;

                // Parse the libraries
                let libs = this._loadOneDrive ? resp["value"] : resp.results;
                Helper.Executor(libs, lib => {
                    // See if we are stopping this process
                    if (this._stopFl) { return; }

                    // Show a dialog
                    this._elSubNav.children[0].innerHTML = `${siteText} - [Analyzing Library ${++ctrList} of ${libs.length}]: ${lib.Title}`;

                    // Analyze the library
                    return this.analyzeLibrary(web, lib);
                }).then(() => {
                    // Resolve the request
                    resolve(null);
                });
            });
        });
    }

    // Gets the form fields to display
    static getFormFields(): Components.IFormControlProps[] {
        return [];
    }

    // Renders the search summary
    private static renderSummary(el: HTMLElement, auditOnly: boolean, items: IAgentItem[], onClose?: () => void) {
        // Render the summary
        this._dashboard = new Dashboard({
            el,
            navigation: {
                title: "Search Agents",
                showFilter: false,
                items: onClose ? [{
                    text: "New Search",
                    className: "btn-outline-light",
                    isButton: true,
                    onClick: () => {
                        // Set the stop flag
                        this._stopFl = true;

                        // Call the close event
                        onClose();
                    }
                }] : null,
                itemsEnd: [{
                    text: "Export to CSV",
                    className: "btn-outline-light me-2",
                    isButton: true,
                    onClick: () => {
                        // Export the CSV
                        new ExportCSV("searchAgents.csv", CSVFields, items);
                    }
                }]
            },
            table: {
                rows: items,
                onRendering: dtProps => {
                    dtProps.columnDefs = [
                        {
                            "targets": 2,
                            "orderable": false,
                            "searchable": false
                        }
                    ];

                    // Order by the 1st column
                    dtProps.order = [[0, "asc"]];

                    // Return the properties
                    return dtProps;
                },
                columns: [
                    {
                        name: "FileName",
                        title: "Object Type"
                    },
                    {
                        name: "ListName",
                        title: "Group Name"
                    },
                    {
                        className: "text-end",
                        name: "",
                        title: "",
                        onRenderCell: (el, col, row: IAgentItem) => {
                            // Render the tooltips
                            let tooltips = Components.TooltipGroup({ el });

                            // See if this is a file
                            if (row.FileUrl) {
                                // Add a button to the file
                                tooltips.add({
                                    content: "Click to view the file.",
                                    btnProps: {
                                        className: "pe-2 py-1",
                                        text: "View File",
                                        type: Components.ButtonTypes.OutlinePrimary,
                                        onClick: () => {
                                            // View the file
                                            window.open(Documents.isWopi(row.FileName) ? row.WebUrl + "/_layouts/15/WopiFrame.aspx?sourcedoc=" + row.FileUrl + "&action=view" : row.FileUrl, "_blank");
                                        }
                                    }
                                });
                            }

                            // Add the view button
                            tooltips.add({
                                content: "Click to view the item unique permissions.",
                                btnProps: {
                                    className: "pe-2 py-1",
                                    //iconType: GetIcon(24, 24, "PeopleTeam", "mx-1"),
                                    text: "View Library",
                                    type: Components.ButtonTypes.OutlinePrimary,
                                    onClick: () => {
                                        // View the library
                                        window.open(row.FileUrl);
                                    }
                                }
                            });
                        }
                    }
                ]
            }
        });

        // Set the sub-nav element
        this._elSubNav = el.querySelector("#sub-navigation");
        this._elSubNav.classList.remove("d-none");
        this._elSubNav.classList.add("my-2");
        this._elSubNav.innerHTML = `<div class="h6"></div><div></div>`;
    }

    // Runs the report
    static run(el: HTMLElement, auditOnly: boolean, values: { [key: string]: any }, onClose: () => void) {
        this._loadOneDrive = values["LoadOneDrive"] == "true";
        this._stopFl = false;

        // Show a loading dialog
        LoadingDialog.setHeader("Searching Sites");
        LoadingDialog.setBody("Loading the sites to search...");
        LoadingDialog.show();

        // Clear the items
        this._items = [];

        // Clear the element
        while (el.firstChild) { el.removeChild(el.firstChild); }

        // Render the summary
        this.renderSummary(el, auditOnly, this._items, onClose);

        // Hide the loading dialog
        LoadingDialog.hide();

        // Determine the webs to target
        let siteItems: Components.IDropdownItem[] = null;
        if (this._loadOneDrive) {
            siteItems = [{ text: DataSource.OneDriveWeb.Url, value: DataSource.OneDriveWeb.Id }] as any;
        } else {
            siteItems = values["TargetWeb"] && values["TargetWeb"]["value"] ? [values["TargetWeb"]] as any : DataSource.SiteItems;
        }

        // Parse all webs
        let counter = 0;
        Helper.Executor(siteItems, siteItem => {
            // See if we are stopping this process
            if (this._stopFl) { return; }

            // Update the status
            this._elSubNav.children[0].innerHTML = `Searching Site ${++counter} of ${siteItems.length}`;
            this._elSubNav.children[1].innerHTML = "Getting the info for the web...";

            // Return a promise
            return new Promise(resolve => {
                // Get the permissions
                let web = this._loadOneDrive ? Web.getOneDrive() : Web(siteItem.text, { requestDigest: DataSource.SiteContext.FormDigestValue });
                web.query({ Select: ["Id", "Title", "Url"] }).execute(web => {
                    // Update the dialog
                    this._elSubNav.children[1].innerHTML = `Analyzing web ${counter} of ${siteItems.length}...`;

                    // Analyze the site
                    this.analyzeSite(web).then(resolve);
                });
            });
        }).then(() => {
            // Hide the sub-nav
            this._elSubNav.classList.add("d-none");
        });
    }

    // Searches a library for agents
    static searchLibrary(webUrl: string, listName: string, auditOnly: boolean) {
        this._loadOneDrive = false;
        this._stopFl = false;

        // Clear the items
        this._items = [];

        // Clear the modal
        Modal.clear();
        Modal.setType(Components.ModalTypes.Full);
        Modal.setHeader("Search Agents Report");
        Modal.setCloseEvent(() => {
            // Set the flag
            this.stop();
        });

        // Render the footer
        Components.ButtonGroup({
            el: Modal.FooterElement,
            buttons: [
                {
                    text: "Close",
                    type: Components.ButtonTypes.OutlinePrimary,
                    onClick: () => {
                        // Set the flag
                        this.stop();
                        Modal.hide();
                    }
                }
            ]
        });

        // Render the summary
        this.renderSummary(Modal.BodyElement, auditOnly, this._items);

        // Show the modal
        Modal.show();

        // Update the status
        this._elSubNav.children[0].innerHTML = `Searching Library: ${listName}`;
        this._elSubNav.children[1].innerHTML = "Getting the info for the web...";

        // Get the permissions
        let web = Web(webUrl, { requestDigest: DataSource.SiteContext.FormDigestValue });
        web.query({ Select: ["Id", "Title", "Url"] }).execute(webInfo => {
            // Update the dialog
            this._elSubNav.children[1].innerHTML = `Analyzing the library...`;

            // Get the list information
            web.Lists(listName).query({
                Expand: ["RootFolder"],
                Select: ["Id", "Title", "BaseTemplate", "HasUniqueRoleAssignments", "RootFolder/ServerRelativeUrl"]
            }).execute(list => {
                // Analyze the library
                this.analyzeLibrary(webInfo, list).then(() => {
                    // Hide the sub-nav
                    this._elSubNav.classList.add("d-none");
                });
            });
        });
    }

    // Stops the report
    static stop() { this._stopFl = true; }
}