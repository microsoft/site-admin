import { CanvasForm, Dashboard, Documents, LoadingDialog, Modal } from "dattatable";
import { Components, Helper, SPTypes, Types, Web } from "gd-sprest-bs";
import { fileEarmark } from "gd-sprest-bs/build/icons/svgs/fileEarmark";
import * as moment from "moment";
import { DataSource } from "../ds";
import Strings from "../strings";
import { ExportCSV } from "./exportCSV";

interface IDLPItem {
    AppliedActionsText: string;
    Author: string;
    ConditionDescription: string;
    FileExtension: string;
    FileName: string;
    GeneralText: string;
    HasUniquePermissions: boolean;
    Id: number;
    LastProcessedTime: string;
    ListId: string;
    ListTitle: string;
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
    "ConditionDescription",
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
                            if (fileExtensions) {
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
                                        // Parse the conditions
                                        result.GetDlpPolicyTip.MatchedConditionDescriptions.results.forEach(condition => {
                                            let dataItem: IDLPItem = {
                                                AppliedActionsText: result.GetDlpPolicyTip.AppliedActionsText,
                                                Author: item["Author"]?.Title,
                                                ConditionDescription: condition,
                                                FileExtension: item["File_x0020_Type"],
                                                FileName: item["FileLeafRef"],
                                                GeneralText: result.GetDlpPolicyTip.GeneralText,
                                                HasUniquePermissions: result.HasUniqueRoleAssignments,
                                                Id: item.Id,
                                                LastProcessedTime: result.GetDlpPolicyTip.LastProcessedTime,
                                                ListId: lib.Id,
                                                ListTitle: lib.Title,
                                                Path: item["FileRef"],
                                                Permissions: result.RoleAssignments.results as any,
                                                WebId: webId,
                                                WebUrl: webUrl
                                            };

                                            // Append the data
                                            this._items.push(dataItem);
                                            this._dashboard.Datatable.addRow(dataItem);
                                        });
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
                    dtProps.columnDefs = [
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
                        name: "ConditionDescription",
                        title: "Condition"
                    },
                    {
                        name: "LastProcessedTime",
                        title: "Last Processed Time",
                        onRenderCell: (el, col, item: IDLPItem) => {
                            el.innerHTML = item.LastProcessedTime ? moment(item.LastProcessedTime).format(Strings.TimeFormat) : "";
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
                                        let groupId = DataSource.getGroupId(role.Member.LoginName);
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
                        title: "",
                        onRenderCell: (el, col, row: IDLPItem) => {
                            let btnDelete: Components.IButton = null;

                            // Set the tooltips
                            let tooltips: Components.ITooltipProps[] = [
                                {
                                    content: "Click to view the document.",
                                    btnProps: {
                                        className: "pe-2 py-1",
                                        iconClassName: "mx-1",
                                        iconType: fileEarmark,
                                        iconSize: 24,
                                        text: "View",
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
                                            this.viewPermissions(row);
                                        }
                                    }
                                }
                            ];

                            // See if the file has broken inheritance
                            if (row.HasUniquePermissions) {
                                // Add the option to revert the permissions
                                tooltips.push({
                                    content: "Click to restore permissions to inherit.",
                                    btnProps: {
                                        assignTo: btn => { btnDelete = btn; },
                                        className: "pe-2 py-1",
                                        text: "Restore",
                                        type: Components.ButtonTypes.OutlinePrimary,
                                        onClick: () => {
                                            // Confirm the deletion of the group
                                            if (confirm("Are you sure you restore the permissions to inherit?")) {
                                                // Revert the permissions
                                                this.revertPermissions(row);
                                            }
                                        }
                                    }
                                });
                            }

                            // Render the buttons
                            Components.TooltipGroup({
                                el,
                                tooltips
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
    private static revertPermissions(item: IDLPItem) {
        // Show a loading dialog
        LoadingDialog.setHeader("Restoring Permissions");
        LoadingDialog.setBody("This window will close after the item permissions are restored...");
        LoadingDialog.show();

        // Restore the permissions
        Web(item.WebUrl, { requestDigest: DataSource.SiteContext.FormDigestValue })
            .Lists().getById(item.ListId).Items(item.Id).resetRoleInheritance().execute(() => {
                // Close the loading dialog
                LoadingDialog.hide();
            });
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
    static searchLibrary(webId: string, webUrl: string, libId: string, libTitle: string) {
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
        this.renderSummary(Modal.BodyElement, false, false);

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
                        // Parse the conditions
                        result.GetDlpPolicyTip.MatchedConditionDescriptions.results.forEach(condition => {
                            let dataItem: IDLPItem = {
                                AppliedActionsText: result.GetDlpPolicyTip.AppliedActionsText,
                                Author: item["Author"]?.Title,
                                ConditionDescription: condition,
                                FileExtension: item["File_x0020_Type"],
                                FileName: item["FileLeafRef"],
                                GeneralText: result.GetDlpPolicyTip.GeneralText,
                                HasUniquePermissions: result.HasUniqueRoleAssignments,
                                Id: item.Id,
                                LastProcessedTime: result.GetDlpPolicyTip.LastProcessedTime,
                                ListId: libId,
                                ListTitle: libTitle,
                                Path: item["FileRef"],
                                Permissions: result.RoleAssignments.results as any,
                                WebId: webId,
                                WebUrl: webUrl
                            };

                            // Append the data
                            this._items.push(dataItem);
                            this._dashboard.Datatable.addRow(dataItem);
                        });
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

    // Views the permissions for the item
    private static viewPermissions(item: IDLPItem) {
        // Clear the canvas form
        CanvasForm.clear();
        CanvasForm.setHeader("View Permissions");
        CanvasForm.setSize(Components.OffcanvasSize.Large2);
        CanvasForm.setType(Components.OffcanvasTypes.End);

        // Set the row data
        let rows: {
            GroupName?: string;
            Member: string;
            Permission: string;
            Type: string;
        }[] = [];

        // Parse the permissions
        item.Permissions.forEach(role => {
            // Parse the role definitions
            let roleDefinitions = [];
            role.RoleDefinitionBindings.results.forEach(roleDef => {
                roleDefinitions.push(`<span><b>${roleDef.Name}:</b> ${roleDef.Description}${roleDef.Hidden ? " (Hidden)" : ""}</span>`);
            });

            // See if this is a user
            switch (role.Member.PrincipalType) {
                case SPTypes.PrincipalTypes.User:
                    rows.push({
                        Member: role.Member.Title || role.Member.LoginName,
                        Type: "User",
                        Permission: roleDefinitions.join('<br/>')
                    });
                    break;
                case SPTypes.PrincipalTypes.SharePointGroup:
                    // Parse the members
                    (role.Member["Users"] ? role.Member["Users"].results : []).forEach(user => {
                        rows.push({
                            Member: user.Email || user.LoginName,
                            Type: "Site Group",
                            GroupName: role.Member.Title,
                            Permission: roleDefinitions.join('<br/>')
                        });
                    });
                    break;
                default:
                    let groupId = DataSource.getGroupId(role.Member.LoginName);
                    rows.push({
                        Member: role.Member.Title,
                        Type: groupId ? "M365 Group" : "AD Group",
                        GroupName: role.Member.Title,
                        Permission: roleDefinitions.join('<br/>')
                    });
                    break;
            }
        });

        // Render the permissions
        new Dashboard({
            el: CanvasForm.BodyElement,
            navigation: {
                showFilter: false,
                title: item.FileName
            },
            table: {
                rows,
                columns: [
                    {
                        name: "Member",
                        title: "Member"
                    },
                    {
                        name: "Type",
                        title: "Type",
                    },
                    {
                        name: "GroupName",
                        title: "Group Name"
                    },
                    {
                        name: "Permission",
                        title: "Permission"
                    }
                ]
            }
        });

        // Show the form
        CanvasForm.show();
    }
}