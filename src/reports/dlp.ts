import { CanvasForm, Dashboard, Documents, LoadingDialog, Modal } from "dattatable";
import { Components, Helper, SPTypes, Types, Web } from "gd-sprest-bs";
import { DataSource } from "../ds";
import { M365Groups } from "../m365Groups";
import Strings from "../strings";
import { ExportCSV } from "./exportCSV";
import { ViewPermissions } from "./viewPermissions";

export interface IDLPItem {
    AppliedActionsText: string;
    Author: string;
    Conditions: string[];
    FileExtension: string;
    FileName: string;
    GeneralText: string;
    HasUniquePermissions: boolean;
    ItemId: number;
    LastProcessedTime: string;
    ListId: string;
    ListTitle: string;
    Overshared: string;
    Path: string;
    Permissions: Types.SP.RoleAssignmentOData[];
    WebUrl: string;
    WebId: string;
}

interface IWebItem {
    ListId: string;
    ListTitle: string;
    WebId: string;
    WebUrl: string;
}

const CSVFields = [
    "WebUrl",
    "ListTitle",
    "FileName",
    "FileExtension",
    "GeneralText",
    "AppliedActionsText",
    "Conditions",
    "LastProcessedTime",
    "Author",
    "Path",
    "ListId",
    "WebId"
]

export class DLP {
    private static _dashboard: Dashboard = null;
    private static _elSubNav: HTMLElement = null;
    private static _items: IDLPItem[] = [];
    private static _loadOneDrive: boolean = false;
    private static _stopFl: boolean = false;

    // Gets the form fields to display
    static getFormFields(fileExt: string = ""): Components.IFormControlProps[] {
        return [
            {
                label: "File Types",
                name: "FileTypes",
                className: "mb-3",
                type: Components.FormControlTypes.TextField,
                value: fileExt
            }
        ];
    }

    // Analyzes the libraries
    private static analyzeLibraries(webId: string, webUrl: string, libraries: Types.SP.ListOData[], fileExtensions: string[]) {
        // Return a promise
        return new Promise(resolve => {
            let counter = 0;
            let siteText = this._elSubNav.children[0].innerHTML;

            // Parse the libraries
            Helper.Executor(libraries, lib => {
                // See if we are stopping this process
                if (this._stopFl) { return; }

                // Update the dialog
                this._elSubNav.children[0].innerHTML = `${siteText} [Analyzing Library ${++counter} of ${libraries.length}]: ${lib.Title}`;

                // Return a promise
                return new Promise(resolve => {
                    let batchRequests = 0;
                    let completed = 0;

                    // Set the list
                    let web = this._loadOneDrive ? Web.getOneDrive() : Web(webUrl, { requestDigest: DataSource.SiteContext.FormDigestValue });
                    let list = web.Lists(lib.Title);

                    // Get the item ids for this library
                    let itemCounter = 0;
                    DataSource.loadItems({
                        isOnedrive: this._loadOneDrive,
                        listId: lib.Id,
                        webUrl,
                        query: {
                            Expand: ["Author"],
                            Select: ["Author/Title", "FileLeafRef", "FileRef", "File_x0020_Type", "Id"]
                        },
                        onItem: item => {
                            let analyzeFile = true;

                            // See if the file extensions are provided
                            if (fileExtensions && fileExtensions.length > 0) {
                                // Default the flag
                                analyzeFile = false

                                // Loop through the file extensions
                                fileExtensions.forEach(fileExt => {
                                    // Set the flag if there is match
                                    if (fileExt.toLowerCase() == item["File_x0020_Type"]?.toLowerCase()) { analyzeFile = true; }
                                });
                            }

                            // See if we are analyzing this file
                            if (analyzeFile) {
                                // Create a batch request to get the dlp policy on this item
                                list.Items(item.Id).query({
                                    Expand: ["GetDlpPolicyTip", "RoleAssignments/Member/Users", "RoleAssignments/RoleDefinitionBindings"],
                                    Select: ["Id", "HasUniqueRoleAssignments"]
                                }).batch(result => {
                                    // Ensure a policy exists
                                    if (result.GetDlpPolicyTip?.MatchedConditionDescriptions) {
                                        let dataItem: IDLPItem = {
                                            AppliedActionsText: result.GetDlpPolicyTip.AppliedActionsText,
                                            Author: item["Author"]?.Title,
                                            Conditions: result.GetDlpPolicyTip.MatchedConditionDescriptions.results,
                                            FileExtension: item["File_x0020_Type"],
                                            FileName: item["FileLeafRef"],
                                            GeneralText: result.GetDlpPolicyTip.GeneralText,
                                            HasUniquePermissions: result.HasUniqueRoleAssignments,
                                            ItemId: item.Id,
                                            LastProcessedTime: result.GetDlpPolicyTip.LastProcessedTime,
                                            ListId: lib.Id,
                                            ListTitle: lib.Title,
                                            Overshared: ViewPermissions.isOvershared(result.RoleAssignments.results as any) ? "Yes" : "No",
                                            Path: item["FileRef"],
                                            Permissions: result.RoleAssignments.results as any,
                                            WebId: webId,
                                            WebUrl: webUrl
                                        };

                                        // Append the data
                                        this._items.push(dataItem);
                                        this._dashboard.Datatable.addRow(dataItem);
                                    }

                                    // Increment the counter and update the dialog
                                    this._elSubNav.children[1].innerHTML = `Batch Requests Processed ${++completed} of ${batchRequests}...`;
                                }, batchRequests++ % Strings.MaxBatchSize == 0);
                            }

                            // Update the dialog
                            this._elSubNav.children[1].innerHTML = `Creating Batch Requests - Processed ${++itemCounter} items...`;

                            // Return the stop flag
                            return this._stopFl;
                        }
                    }).then(() => {
                        // Update the dialog
                        this._elSubNav.children[1].innerHTML = `Executing Batch Request for ${batchRequests} items...`;

                        // Execute the batch request
                        list.execute(resolve);
                    }, resolve);
                });
            }).then(resolve);
        });
    }

    // Refreshes the item
    private static refreshItem(dlpItem: IDLPItem): PromiseLike<IDLPItem> {
        // Return a promise
        return new Promise(resolve => {
            let web = this._loadOneDrive ? Web.getOneDrive() : Web(dlpItem.WebUrl, { requestDigest: DataSource.SiteContext.FormDigestValue });

            // Get the list item
            web.Lists(dlpItem.ListTitle).Items(dlpItem.ItemId).query({
                Expand: ["GetDlpPolicyTip", "RoleAssignments/Member/Users", "RoleAssignments/RoleDefinitionBindings"],
                Select: ["Id", "HasUniqueRoleAssignments"]
            }).execute((item) => {
                // Update the item
                dlpItem.HasUniquePermissions = item.HasUniqueRoleAssignments;
                dlpItem.Overshared = ViewPermissions.isOvershared(item.RoleAssignments.results as any) ? "Yes" : "No";
                dlpItem.Permissions = item.RoleAssignments.results as any;

                // Ensure a policy exists
                if (item.GetDlpPolicyTip?.MatchedConditionDescriptions) {
                    // Update the item
                    dlpItem.AppliedActionsText = item.GetDlpPolicyTip.AppliedActionsText;
                    dlpItem.Conditions = item.GetDlpPolicyTip.MatchedConditionDescriptions.results;
                    dlpItem.GeneralText = item.GetDlpPolicyTip.GeneralText;
                    dlpItem.LastProcessedTime = item.GetDlpPolicyTip.LastProcessedTime;
                } else {
                    // Update the item
                    dlpItem.AppliedActionsText = "";
                    dlpItem.Conditions = [];
                    dlpItem.GeneralText = "";
                    dlpItem.LastProcessedTime = "";
                }

                // Find the item
                for (let i = 0; i < this._items.length; i++) {
                    let item = this._items[i];

                    // See if this is the item
                    if (item.ItemId == dlpItem.ItemId && item.ListId == dlpItem.ListId) {
                        // Update the item
                        this._items[i] = dlpItem;
                    }
                }

                // Resolve the request
                resolve(dlpItem);
            });
        });
    }

    // Renders the search summary
    private static renderSummary(el: HTMLElement, auditOnly: boolean, showSearch?: boolean, onClose?: () => void) {
        // Render the summary
        this._dashboard = new Dashboard({
            el,
            navigation: {
                title: "DLP Report",
                showFilter: false,
                items: showSearch ? [{
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
                        new ExportCSV("dlpReport.csv", CSVFields, this._items);
                    }
                }]
            },
            table: {
                rows: this._items,
                onRendering: dtProps => {
                    // Set the column defs
                    dtProps.columnDefs = [
                        {
                            "targets": [2],
                            "orderable": false
                        },
                        {
                            "targets": [4, 5],
                            "orderable": false,
                            "searchable": false
                        }
                    ];

                    // Order by the 2nd column by default; ascending
                    dtProps.order = [[1, "asc"]];

                    // Return the properties
                    return dtProps;
                },
                columns: [
                    {
                        name: "ListTitle",
                        title: "List"
                    },
                    {
                        name: "Path",
                        title: "File",
                        onRenderCell: (el, col, item: IDLPItem) => {
                            // Set the order info
                            el.setAttribute("data-order", item.FileName);

                            // Show the file info
                            el.innerHTML = `
                                <b>Name: </b>${item.FileName}
                                <br/>
                                <b>Created By: </b>${item.Author}
                                <br/>
                                <b>Path: </b>${item.Path}
                            `;
                        }
                    },
                    {
                        name: "",
                        title: "Condition(s)",
                        onRenderCell: (el, col, item: IDLPItem) => {
                            // Render the conditions
                            let elList = document.createElement("ul");
                            item.Conditions.forEach(condition => {
                                // Create the item
                                let elItem = document.createElement("li");
                                elItem.innerText = condition;
                                elList.appendChild(elItem);
                            });

                            // Append the list to the cell
                            el.appendChild(elList);
                        }
                    },
                    {
                        name: "",
                        title: "Overshared",
                        onRenderCell: (el, col, item: IDLPItem) => {
                            let isOvershared = item.Overshared === "Yes" ? true : false;

                            // Set the order info
                            el.setAttribute("data-order", item.Overshared);

                            // Make the badge display in the middle
                            el.style.verticalAlign = "middle";

                            // Render a badge
                            let badge = Components.Badge({
                                el,
                                className: "me-2",
                                content: isOvershared ? "Overshared" : item.Overshared,
                                type: isOvershared ? Components.BadgeTypes.Danger : Components.BadgeTypes.Secondary,
                                isPill: true
                            });

                            // See if this is overshared
                            if (isOvershared) {
                                // Render a tooltip
                                Components.Tooltip({
                                    target: badge.el,
                                    content: `The file has been flagged as overshared because it's shared with the following groups:<br/>${ViewPermissions.getOversharedGroups(item.Permissions).join("<br/>")}`
                                });
                            }
                        }
                    },
                    {
                        name: "",
                        title: "Permissions",
                        onRenderCell: (el, col, item: IDLPItem) => {
                            let adGroups = 0;
                            let m365Groups = 0;
                            let siteGroups = 0;
                            let users = 0;

                            // Parse the permissions
                            item.Permissions.forEach(role => {
                                // See if this is a user
                                switch (role.Member.PrincipalType) {
                                    case SPTypes.PrincipalTypes.User:
                                        users++;
                                        break;
                                    case SPTypes.PrincipalTypes.SharePointGroup:
                                        siteGroups++;
                                        break;
                                    default:
                                        let groupId = M365Groups.getGroupId(role.Member.LoginName);
                                        groupId ? m365Groups++ : adGroups++;
                                        break;
                                }
                            });

                            // Output the permission information
                            el.innerHTML = `
                                <b>Unique Permissions: </b>${item.HasUniquePermissions ? "Yes" : "No"}
                                <br/>
                                <b># of Users: </b>${users}
                                <br/>
                                <b># of Site Groups: </b>${siteGroups}
                                <br/>
                                <b># of AD Groups: </b>${adGroups}
                                <br/>
                                <b># of M365 Groups: </b>${m365Groups}
                                <br/>
                            `;
                        }
                    },
                    {
                        className: "text-end",
                        name: "",
                        title: "Actions",
                        onRenderCell: (el, col, row: IDLPItem, rowIdx) => {
                            // Set the tooltips
                            let tooltips: Components.ITooltipProps[] = [
                                {
                                    content: "Click to view the document.",
                                    btnProps: {
                                        className: "pe-2 py-1",
                                        text: "View File",
                                        type: Components.ButtonTypes.OutlinePrimary,
                                        onClick: () => {
                                            // View the document
                                            window.open(Documents.isWopi(`${row.FileName}`) ? row.WebUrl + "/_layouts/15/WopiFrame.aspx?sourcedoc=" + row.Path + "&action=view" : row.Path, "_blank");
                                        }
                                    }
                                },
                                {
                                    content: "Click to view the permissions for this document.",
                                    btnProps: {
                                        className: "pe-2 py-1",
                                        text: "View Permissions",
                                        type: Components.ButtonTypes.OutlinePrimary,
                                        onClick: () => {
                                            // View the permissions for the document
                                            ViewPermissions.show(row);
                                        }
                                    }
                                }
                            ];

                            // See if the file is overshared
                            if (!auditOnly && row.Overshared === "Yes") {
                                tooltips.push({
                                    content: "Click to remove the groups that are flagging this file as overshared.",
                                    btnProps: {
                                        className: "pe-2 py-1",
                                        text: "Secure File",
                                        type: Components.ButtonTypes.OutlinePrimary,
                                        onClick: () => {
                                            // Remove the overshared groups from the permissions
                                            ViewPermissions.removeOversharedGroups(row, () => {
                                                // Refresh the item
                                                this.refreshItem(row).then(updatedItem => {
                                                    // Update this row
                                                    this._dashboard.updateRow(rowIdx, updatedItem);
                                                });
                                            });
                                        }
                                    }
                                });
                            }

                            // See if the file has broken inheritance
                            if (!auditOnly && row.HasUniquePermissions) {
                                // Add the option to revert the permissions
                                tooltips.push({
                                    content: "Click to restore permissions to inherit.",
                                    btnProps: {
                                        className: "pe-2 py-1",
                                        text: "Restore Permissions",
                                        type: Components.ButtonTypes.OutlinePrimary,
                                        onClick: () => {
                                            // Revert the permissions
                                            this.revertPermissions(row, () => {
                                                // Refresh the item
                                                this.refreshItem(row).then(updatedItem => {
                                                    // Update this row
                                                    this._dashboard.updateRow(rowIdx, updatedItem);
                                                });
                                            });
                                        }
                                    }
                                });
                            }

                            // Render the buttons
                            Components.TooltipGroup({
                                el,
                                tooltips,
                                isSmall: true,
                                isVertical: true
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

    // Reverts the item permissions
    private static revertPermissions(item: IDLPItem, onComplete: () => void) {
        // Clear a canvas form
        CanvasForm.clear();
        CanvasForm.setHeader("Secure File");
        CanvasForm.setType(Components.OffcanvasTypes.End);
        CanvasForm.setSize(Components.OffcanvasSize.Small2);

        // Set the content
        CanvasForm.setBody(`
            <p>Are you sure you want to revert the permissions to inherit from its parent? Click on 'Confirm' to complete the request.</p>
            <div class="d-flex justify-content-end"></div>
        `);

        // Render the footer
        Components.ButtonGroup({
            el: CanvasForm.BodyElement.querySelector("div"),
            className: "mt-3",
            buttons: [
                {
                    text: "Confirm",
                    type: Components.ButtonTypes.OutlinePrimary,
                    onClick: () => {
                        // Show a loading dialog
                        LoadingDialog.setHeader("Restoring Permissions");
                        LoadingDialog.setBody("This window will close after the item permissions are restored...");
                        LoadingDialog.show();

                        // Restore the permissions
                        Web(item.WebUrl, { requestDigest: DataSource.SiteContext.FormDigestValue })
                            .Lists().getById(item.ListId).Items(item.ItemId).resetRoleInheritance().execute(() => {
                                // Hide the dialogs
                                LoadingDialog.hide();
                                CanvasForm.hide();

                                // Call the event
                                onComplete();
                            });
                    }
                },
                {
                    text: "Close",
                    type: Components.ButtonTypes.OutlineSecondary,
                    onClick: () => {
                        // Hide the form
                        CanvasForm.hide();
                    }
                }
            ]
        });

        // Show the canvas form
        CanvasForm.show();
    }

    // Runs the report
    static run(el: HTMLElement, auditOnly: boolean, values: { [key: string]: string }, onClose: () => void) {
        let data: IWebItem[] = [];
        this._loadOneDrive = values["LoadOneDrive"] == "true";
        this._stopFl = false;

        // Clear the items
        this._items = [];

        // Show a loading dialog
        LoadingDialog.setHeader("Searching Site");
        LoadingDialog.setBody("Loading the libraries...");
        LoadingDialog.show();

        // Get the file extensions
        let fileExtensions: string[] = values["FileTypes"] ? values["FileTypes"].trim().split(' ') : [];

        // Clear the element
        while (el.firstChild) { el.removeChild(el.firstChild); }

        // Render the summary
        this.renderSummary(el, auditOnly, true, onClose);

        // Hide the loading dialog
        LoadingDialog.hide();

        // Determine the webs to target
        let siteItems: Components.IDropdownItem[] = null;
        if (this._loadOneDrive) {
            siteItems = [{ text: DataSource.OneDriveWeb.Url, value: DataSource.OneDriveWeb.Id }] as any;
        } else {
            siteItems = values["TargetWeb"] && values["TargetWeb"]["value"] ? [values["TargetWeb"]] as any : DataSource.SiteItems;
        }

        // Parse the webs
        let counter = 0;
        Helper.Executor(siteItems, siteItem => {
            // See if we are stopping this process
            if (this._stopFl) { return; }

            // Update the status
            this._elSubNav.children[0].innerHTML = `Searching Site ${++counter} of ${siteItems.length}`;

            // Return a promise
            return new Promise(resolve => {
                // Get the libraries for this site
                let web = this._loadOneDrive ? Web.getOneDrive() : Web(siteItem.text, { requestDigest: DataSource.SiteContext.FormDigestValue });
                web.Lists().query({
                    Filter: `Hidden eq false and BaseTemplate eq ${SPTypes.ListTemplateType.DocumentLibrary} or BaseTemplate eq ${SPTypes.ListTemplateType.MySiteDocumentLibrary} or BaseTemplate eq ${SPTypes.ListTemplateType.PageLibrary}`,
                    GetAllItems: true,
                    Select: ["Id", "Title"],
                    Top: 5000
                }).execute(libs => {
                    // Add the libraries to analyze
                    libs.results.forEach(lib => {
                        // Append the item
                        data.push({
                            ListId: lib.Id,
                            ListTitle: lib.Title,
                            WebId: siteItem.value,
                            WebUrl: siteItem.text
                        });
                    });

                    // Update the dialog
                    this._elSubNav.children[1].innerHTML = "Loading the files for the libraries...";

                    // Analyze the libraries
                    this.analyzeLibraries(siteItem.value, siteItem.text, libs.results, fileExtensions).then(resolve);
                });
            });
        }).then(() => {
            // Hide the sub-nav
            this._elSubNav.classList.add("d-none");
        });
    }

    // Searches a library for DLP conditions
    static searchLibrary(webId: string, webUrl: string, libId: string, libTitle: string, auditOnly: boolean) {
        // Clear the items
        this._items = [];

        // Set the modal header
        Modal.clear();
        Modal.setType(Components.ModalTypes.Full);
        Modal.setHeader("Data Loss Prevention Report");
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

        // Show the results
        this.renderSummary(Modal.BodyElement, auditOnly, false);

        // Show the modal
        Modal.show();

        // Update the status
        this._elSubNav.children[0].innerHTML = "Analyzing Library";
        this._elSubNav.children[1].innerHTML = "Getting all files in this library...";

        // Create the list for the batch requests
        let batchRequests = 0;
        let completed = 0;
        let web = this._loadOneDrive ? Web.getOneDrive() : Web(webUrl, { requestDigest: DataSource.SiteContext.FormDigestValue });
        let list = web.Lists().getById(libId);

        // Get the item ids for this library
        let itemCounter = 0;
        DataSource.loadItems({
            webUrl,
            listId: libId,
            query: {
                Expand: ["Author"],
                Select: ["Author/Title", "FileLeafRef", "FileRef", "File_x0020_Type", "Id"],
            },
            onItem: item => {
                // Create a batch request to get the dlp policy on this item
                list.Items(item.Id).query({
                    Expand: ["GetDlpPolicyTip", "RoleAssignments/Member/Users", "RoleAssignments/RoleDefinitionBindings"],
                    Select: ["Id", "HasUniqueRoleAssignments"]
                }).batch(result => {
                    // Ensure a policy exists
                    if (result.GetDlpPolicyTip?.MatchedConditionDescriptions) {
                        // Create the item
                        let dataItem: IDLPItem = {
                            AppliedActionsText: result.GetDlpPolicyTip.AppliedActionsText,
                            Author: item["Author"]?.Title,
                            Conditions: result.GetDlpPolicyTip.MatchedConditionDescriptions.results,
                            FileExtension: item["File_x0020_Type"],
                            FileName: item["FileLeafRef"],
                            GeneralText: result.GetDlpPolicyTip.GeneralText,
                            HasUniquePermissions: result.HasUniqueRoleAssignments,
                            ItemId: item.Id,
                            LastProcessedTime: result.GetDlpPolicyTip.LastProcessedTime,
                            ListId: libId,
                            ListTitle: libTitle,
                            Overshared: ViewPermissions.isOvershared(result.RoleAssignments.results as any) ? "Yes" : "No",
                            Path: item["FileRef"],
                            Permissions: result.RoleAssignments.results as any,
                            WebId: webId,
                            WebUrl: webUrl
                        };

                        // Append the data
                        this._items.push(dataItem);
                        this._dashboard.Datatable.addRow(dataItem);
                    }

                    // Increment the counter and update the dialog
                    this._elSubNav.children[1].innerHTML = `Batch Requests Processed ${++completed} of ${batchRequests}...`;
                }, batchRequests++ % Strings.MaxBatchSize == 0);

                // Update the dialog
                this._elSubNav.children[1].innerHTML = `Creating Batch Requests - Processed ${++itemCounter} items...`;

                // Return the stop flag
                return this._stopFl;
            }
        }).then(() => {
            // Update the dialog
            this._elSubNav.children[1].innerHTML = `Executing Batch Request for ${batchRequests} items...`;

            // Execute the batch request
            list.execute(() => {
                // Hide the sub-nav
                this._elSubNav.classList.add("d-none");
            });
        });
    }

    // Stops the report
    static stop() { this._stopFl = true; }
}