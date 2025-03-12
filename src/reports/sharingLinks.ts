import { Dashboard, Documents, LoadingDialog } from "dattatable";
import { Components, ContextInfo, Helper, Search, Types, Web } from "gd-sprest-bs";
import { fileEarmark } from "gd-sprest-bs/build/icons/svgs/fileEarmark";
import { personX } from "gd-sprest-bs/build/icons/svgs/personX";
import { DataSource } from "../ds";
import { ExportCSV } from "./exportCSV";

interface IGroupInfo {
    FileExtension?: string;
    FileName?: string;
    FileUrl?: string;
    Role?: string;
    RoleInfo?: string;
    Name: string;
    Email?: string;
    Group?: string;
    GroupId?: number;
    GroupInfo?: string;
    Id?: number;
    WebUrl: string;
}

interface IUserInfo {
    EMail: string;
    Id: number;
    Name: string;
    Title: string;
}

const CSVFields = [
    "WebUrl", "Name", "Email", "Group", "GroupId", "GroupInfo", "Role", "RoleInfo", "FileName", "FileExtension", "FileUrl"
]

export class SharingLinks {
    private static _items: IGroupInfo[] = [];

    // Gets the associated file for this sharing link
    private static analyzeDocInfo(rootWeb: Types.SP.WebOData, docInfo: { docId: string, group: Types.SP.Group, roleName: string, userInfo: IUserInfo }): PromiseLike<void> {
        // Return a promise
        return new Promise(resolve => {
            // Find the document by its id
            Search.postQuery<{ FileExtension: string; Path: string; SPWebUrl: string; Title: string }>({
                url: DataSource.SiteContext.SiteFullUrl,
                query: {
                    Querytext: "UniqueID: " + docInfo.docId,
                    SelectProperties: {
                        results: ["FileExtension", "Path", "SPWebUrl", "Title"]
                    }
                },
                targetInfo: { requestDigest: DataSource.SiteContext.FormDigestValue }
            }).then(search => {
                let result = search.results[0];

                // Set the role information
                let roleInfo = "";
                if (result) {
                    // Set the role information
                    roleInfo = "Has '" + docInfo.roleName + "' access to the file <a target='_blank' " +
                        "href='" + result.Path + "'>" + result.Title + "</a>.";
                }

                // Add the user information
                this._items.push({
                    FileExtension: result?.Path,
                    FileName: result?.Title,
                    FileUrl: result?.Path,
                    WebUrl: result?.SPWebUrl || rootWeb.Url,
                    Name: docInfo.userInfo.Title || docInfo.userInfo.Name,
                    Email: docInfo.userInfo.EMail,
                    Group: docInfo.group.Title,
                    GroupId: docInfo.group.Id,
                    GroupInfo: docInfo.group.Description,
                    Role: docInfo.roleName,
                    RoleInfo: roleInfo
                });

                // Resolve the promise
                resolve(null);
            });
        });
    }

    // Analyzes the groups
    private static analyzeGroups(rootWeb: Types.SP.WebOData, groups: IUserInfo[]): PromiseLike<void> {
        // Return a promise
        return new Promise(resolve => {
            let docInfo: { docId: string, group: Types.SP.Group, roleName: string, userInfo: IUserInfo }[] = [];
            let web = Web(DataSource.SiteContext.SiteFullUrl, { requestDigest: DataSource.SiteContext.FormDigestValue });

            // Update the loading dialog
            LoadingDialog.setBody(`Creating the batch job to get the group information...`);

            // Get the group ids
            let groupIds: number[] = [];
            let groupMapper: { [key: number]: IUserInfo } = {};
            groups.forEach(group => {
                // Add the group id
                groupIds.push(group.Id);
                groupMapper[group.Id] = group;
            });

            // Parse the group ids
            groupIds.forEach(groupId => {
                // Get the group
                web.SiteGroups().getById(groupId).batch(group => {
                    // Set the role name
                    let info = group.Title.split('.');
                    let docId = info.length > 2 ? info[1] : "";
                    let roleName = info.length > 3 ? info[2] : "";

                    // See if the doc id exists
                    if (docId) {
                        // Add the doc info
                        docInfo.push({ docId, group, roleName, userInfo: groupMapper[group.Id] });
                    } else {
                        let groupInfo = groupMapper[group.Id];

                        // Add the group information
                        this._items.push({
                            WebUrl: rootWeb.Url,
                            Name: groupInfo.Title || groupInfo.Name,
                            Email: groupInfo.EMail,
                            Group: group.Title,
                            GroupId: group.Id,
                            GroupInfo: group.Description,
                            Role: roleName
                        });
                    }
                });
            });

            // Update the loading dialog
            LoadingDialog.setBody(`Executing the batch job for the group information...`);

            // Execute the batch job
            web.execute(() => {
                // Parse the doc information
                let counter = 0;
                Helper.Executor(docInfo, docInfo => {
                    // Update the loading dialog
                    LoadingDialog.setBody(`Analyzing Sharing Link ${++counter} of ${groups.length}`);

                    // Analyze the group
                    return this.analyzeDocInfo(rootWeb, docInfo);
                }).then(() => {
                    // Resolve the request
                    resolve();
                });
            });
        });
    }

    // Gets the form fields to display
    static getFormFields(): Components.IFormControlProps[] { return []; }

    // Removes a group
    private static removeGroup(group: string, groupId: number) {
        // Display a loading dialog
        LoadingDialog.setHeader("Removing Site User");
        LoadingDialog.setBody(`Removing the group '${group}' group. This will close after the request completes.`);
        LoadingDialog.show();

        // Remove the user from the group
        Web(DataSource.SiteContext.SiteFullUrl, { requestDigest: DataSource.SiteContext.FormDigestValue })
            .SiteGroups().removeById(groupId).execute(
                // Success
                () => {
                    // Close the dialog
                    LoadingDialog.hide();
                },

                // Error
                () => {
                    // Close the dialog
                    LoadingDialog.hide();
                }
            )
    }

    // Renders the search summary
    private static renderSummary(el: HTMLElement, onClose: () => void) {
        // Render the summary
        new Dashboard({
            el,
            navigation: {
                title: "Sharing Links",
                showFilter: false,
                items: [{
                    text: "New Search",
                    className: "btn-outline-light",
                    isButton: true,
                    onClick: () => {
                        // Call the close event
                        onClose();
                    }
                }],
                itemsEnd: [{
                    text: "Export to CSV",
                    className: "btn-outline-light me-2",
                    isButton: true,
                    onClick: () => {
                        // Export the CSV
                        new ExportCSV("listPermissions.csv", CSVFields, this._items);
                    }
                }]
            },
            table: {
                rows: this._items,
                onRendering: dtProps => {
                    dtProps.columnDefs = [
                        {
                            "targets": 5,
                            "orderable": false,
                            "searchable": false
                        }
                    ];

                    // Order by the 1st column by default; ascending
                    dtProps.order = [[0, "asc"]];

                    // Return the properties
                    return dtProps;
                },
                columns: [
                    {
                        name: "WebUrl",
                        title: "Url"
                    },
                    {
                        name: "Group",
                        title: "Group Name"
                    },
                    {
                        name: "GroupInfo",
                        title: "Group Detail",
                        onRenderCell: (el) => {
                            // Add the data-filter attribute for searching notes properly
                            el.setAttribute("data-filter", el.innerHTML);
                            // Add the data-order attribute for sorting notes properly
                            el.setAttribute("data-order", el.innerHTML);

                            // Declare a span element
                            let span = document.createElement("span");
                            span.className = "notes";

                            // Return the plain text if less than 50 chars
                            if (el.innerHTML.length < 50) {
                                span.innerHTML = el.innerHTML;
                            } else {
                                // Truncate to the last white space character in the text after 50 chars and add an ellipsis
                                span.innerHTML = el.innerHTML.substring(0, 50).replace(/\s([^\s]*)$/, '') + '&#8230';

                                // Add a tooltip containing the text
                                Components.Tooltip({
                                    content: "<small>" + el.innerHTML + "</small>",
                                    target: span
                                });
                            }

                            // Clear the element
                            el.innerHTML = "";
                            // Append the span
                            el.appendChild(span);
                        }
                    },
                    {
                        name: "Role",
                        title: "Permission"
                    },
                    {
                        name: "RoleInfo",
                        title: "Permission Detail",
                        onRenderCell: (el) => {
                            // Add the data-filter attribute for searching notes properly
                            el.setAttribute("data-filter", el.innerHTML);
                            // Add the data-order attribute for sorting notes properly
                            el.setAttribute("data-order", el.innerHTML);

                            // Declare a span element
                            let span = document.createElement("span");
                            span.className = "notes";

                            // Return the plain text if less than 50 chars
                            if (el.innerHTML.length < 50) {
                                span.innerHTML = el.innerHTML;
                            } else {
                                // Truncate to the last white space character in the text after 50 chars and add an ellipsis
                                span.innerHTML = el.innerHTML.substring(0, 50).replace(/\s([^\s]*)$/, '') + '&#8230';

                                // Add a tooltip containing the text
                                Components.Tooltip({
                                    content: "<small>" + el.innerHTML + "</small>",
                                    target: span
                                });
                            }

                            // Clear the element
                            el.innerHTML = "";

                            // Append the span
                            el.appendChild(span);
                        }
                    },
                    {
                        className: "text-end",
                        name: "",
                        title: "",
                        onRenderCell: (el, col, row: IGroupInfo) => {
                            let btnDelete: Components.IButton = null;
                            let tooltips: Components.ITooltipProps[] = [];

                            // See if a file is associated with this sharing link
                            if (row.FileUrl) {
                                // Add the view file button
                                tooltips.push({
                                    content: "Click to view the file.",
                                    btnProps: {
                                        className: "pe-2 py-1",
                                        iconType: fileEarmark(24, 24, "mx-1"),
                                        text: "File",
                                        type: Components.ButtonTypes.OutlinePrimary,
                                        onClick: () => {
                                            // View the file
                                            window.open(Documents.isWopi(`${row.FileName}.${row.FileExtension}`) ? row.WebUrl + "/_layouts/15/WopiFrame.aspx?sourcedoc=" + row.FileUrl + "&action=view" : row.FileUrl, "_blank");
                                        }
                                    }
                                });
                            }

                            // Add the view group button
                            tooltips.push({
                                content: "Click to view the site group",
                                btnProps: {
                                    className: "pe-2 py-1",
                                    iconType: fileEarmark(24, 24, "mx-1"),
                                    text: "Group",
                                    type: Components.ButtonTypes.OutlinePrimary,
                                    onClick: () => {
                                        // View the group
                                        window.open(`${row.WebUrl}/${ContextInfo.layoutsUrl}/people.aspx?MembershipGroupId=${row.GroupId}`);
                                    }
                                }
                            });

                            // Add the delete group button
                            tooltips.push({
                                el,
                                content: "Click to delete the site group.",
                                btnProps: {
                                    assignTo: btn => { btnDelete = btn; },
                                    className: "pe-2 py-1",
                                    iconType: personX(24, 24, "mx-1"),
                                    text: "Delete",
                                    type: Components.ButtonTypes.OutlineDanger,
                                    onClick: () => {
                                        // Confirm the deletion of the group
                                        if (confirm("Are you sure you want to delete this group?")) {
                                            // Disable this button
                                            btnDelete.disable();

                                            // Delete the site group
                                            this.removeGroup(row.Group, row.GroupId);
                                        }
                                    }
                                }
                            });

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
    }

    // Runs the report
    static run(el: HTMLElement, values: { [key: string]: string }, onClose: () => void) {
        // Show a loading dialog
        LoadingDialog.setHeader("Loading Security Groups");
        LoadingDialog.setBody("Loading the permissions for this site...");
        LoadingDialog.show();

        // Clear the items
        this._items = [];

        let users: IUserInfo[] = [];

        // Load the group information
        Web(DataSource.SiteContext.SiteFullUrl, { requestDigest: DataSource.SiteContext.FormDigestValue }).query({
            Expand: [
                "ParentWeb", "RoleAssignments", "RoleAssignments/Groups", "RoleAssignments/Member",
                "RoleAssignments/Member/Users", "RoleAssignments/RoleDefinitionBindings"
            ]
        }).execute(web => {
            // Update the loading dialog
            LoadingDialog.setBody("Loading the sharing link groups...");

            // Get the users for this site
            Web(DataSource.SiteContext.SiteFullUrl, { requestDigest: DataSource.SiteContext.FormDigestValue }).Lists("User Information List").Items().query({
                Filter: "substringof('SharingLinks', Name)",
                Select: ["Id", "Name", "EMail", "Title"],
                GetAllItems: true,
                Top: 5000
            }).execute(items => {
                // Parse the items
                for (let i = 0; i < items.results.length; i++) {
                    let item = items.results[i];

                    // Add the user
                    users.push({
                        EMail: item["EMail"],
                        Id: item.Id,
                        Name: item["Name"],
                        Title: item.Title
                    });
                }

                // Update the loading dialog
                LoadingDialog.setHeader("Analyzing Sharing Link Groups");

                // Analyze the groups
                this.analyzeGroups(web, users).then(() => {
                    // Clear the element
                    while (el.firstChild) { el.removeChild(el.firstChild); }

                    // Render the summary
                    this.renderSummary(el, onClose);

                    // Hide the loading dialog
                    LoadingDialog.hide();
                });
            });
        });
    }
}