import { Dashboard, Documents, LoadingDialog } from "dattatable";
import { Components, ContextInfo, Helper, SPTypes, Types, Web } from "gd-sprest-bs";
import { DataSource } from "../ds";
import { ExportCSV } from "./exportCSV";

interface ISearchItem {
    Email?: string;
    FileName?: string;
    FileUrl?: string;
    Group?: string;
    GroupId?: number;
    GroupInfo?: string;
    Id?: number;
    ItemId?: number;
    ListId?: string;
    ListName?: string;
    ListUrl?: string;
    LoginName: string;
    Name: string;
    Role?: string;
    RoleInfo?: string;
    WebTitle: string;
    WebUrl: string;
}

interface IUserInfo {
    EMail: string;
    Id: number;
    Name: string;
    Title: string;
    UserName: string;
}

const CSVFields = [
    "Name", "UserName", "Email", "Group", "GroupInfo", "FileName", "FileUrl",
    "ItemId", "ListName", "ListUrl", "Role", "RoleInfo", "WebTitle", "WebUrl"
]

export class SearchEEEU {
    private static _items: ISearchItem[] = null;

    // Analyzes a lists
    private static analyzeList(web: Types.SP.WebOData, list: Types.SP.ListOData): PromiseLike<void> {
        // Return a promise
        return new Promise(resolve => {
            let Select = ["Id", "HasUniqueRoleAssignments"];

            // See if this is a document library
            if (list.BaseTemplate == SPTypes.ListTemplateType.DocumentLibrary || list.BaseTemplate == SPTypes.ListTemplateType.PageLibrary) {
                // Get the file information
                Select.push("FileLeafRef");
                Select.push("FileRef");
            }

            // Get the items where it has broken inheritance
            Web(web.Url, { requestDigest: DataSource.SiteContext.FormDigestValue }).Lists().getById(list.Id).Items().query({
                GetAllItems: true,
                Select,
                Top: 5000
            }).execute(items => {
                // Create a batch job
                let batch = Web(web.Url, { requestDigest: DataSource.SiteContext.FormDigestValue }).Lists().getById(list.Id);

                // Parse the items
                Helper.Executor(items.results, item => {
                    // See if this item doesn't have unique permissions
                    if (!item.HasUniqueRoleAssignments) { return; }

                    // Get the permissions
                    batch.Items(item.Id).RoleAssignments().query({
                        Filter: `Member/Title eq 'Everyone' or substringof('spo-grid-all-users', Member/LoginName)`,
                        Expand: [
                            "Member", "RoleDefinitionBindings"
                        ]
                    }).batch(roles => {
                        // Parse the role assignments
                        Helper.Executor(roles.results, roleAssignment => {
                            let roleDefinition = roleAssignment.RoleDefinitionBindings.results[0];
                            let user: Types.SP.User = roleAssignment.Member as any;

                            // Add a row for this entry
                            this._items.push({
                                Email: user.Email,
                                FileName: item["FileLeafRef"],
                                FileUrl: item["FileRef"],
                                Group: "",
                                GroupId: 0,
                                GroupInfo: "",
                                Id: user.Id,
                                ItemId: item.Id,
                                ListId: list.Id,
                                ListName: list.Title,
                                LoginName: user.LoginName,
                                ListUrl: list.RootFolder.ServerRelativeUrl,
                                Name: user.Title || user.LoginName,
                                Role: roleDefinition.Name,
                                RoleInfo: roleDefinition.Description || "",
                                WebUrl: web.Url,
                                WebTitle: web.Title
                            });
                        });
                    });
                }).then(() => {
                    // Execute the batch job
                    batch.execute(() => {
                        // Resolve the request
                        resolve();
                    });
                });
            });
        });
    }

    // Analyzes the site
    private static analyzeSite(web: Types.SP.WebOData, searchLists: boolean): PromiseLike<void> {
        // Return a promise
        return new Promise(resolve => {
            // Show a loading dialog
            LoadingDialog.setBody("Getting the user information...");

            // Search the users
            this.analyzeUsers(web).then(() => {
                // See if we are searching lists
                if (searchLists) {
                    // Get the lists
                    Web(web.Url, { requestDigest: DataSource.SiteContext.FormDigestValue }).Lists().query({
                        Filter: "Hidden eq false",
                        Expand: ["RootFolder"],
                        Select: ["Id", "Title", "BaseTemplate", "HasUniqueRoleAssignments", "RootFolder/ServerRelativeUrl"]
                    }).execute(lists => {
                        let ctrList = 0;

                        // Parse the lists
                        Helper.Executor(lists.results, list => {
                            // Update the status
                            LoadingDialog.setBody(`Analyzing List ${++ctrList} of ${lists.results.length}...`);

                            // Analyze the list
                            return this.analyzeList(web, list);
                        }).then(() => {
                            // Resolve the request
                            resolve(null);
                        });
                    });
                } else {
                    // Resolve the request
                    resolve();
                }
            })
        });
    }

    // Analyze the site users
    private static analyzeUsers(web: Types.SP.WebOData) {
        return new Promise(resolve => {
            // Get the users
            this.getUsers().then(users => {
                let counter = 0;

                // Parse the users
                Helper.Executor(users, user => {
                    // Update the loading dialog
                    LoadingDialog.setBody(`Analyzing User ${++counter} of ${users.length}`);

                    // Get the user information
                    return this.getUserInfo(web, user);
                }).then(() => {
                    // Resolve the request
                    resolve(null);
                });
            }, resolve);
        })
    }

    // Gets the form fields to display
    static getFormFields(): Components.IFormControlProps[] {
        return [
            {
                name: "SearchLists",
                label: "In Depth Search?",
                description: "Selecting this option will include a search for unique item permissions in lists/libraries.",
                type: Components.FormControlTypes.Switch,
                value: false
            }
        ];
    }

    // Get the user information
    private static getUserInfo(web: Types.SP.WebOData, userInfo: IUserInfo) {
        // Return a promise
        return new Promise((resolve) => {
            // Parse the roles
            for (let i = 0; i < web.RoleAssignments.results.length; i++) {
                let role: Types.SP.RoleAssignmentOData = web.RoleAssignments.results[i] as any;

                // See if this role is the user
                if (role.Member.LoginName == userInfo.Name) {
                    // Add the user information
                    this._items.push({
                        WebUrl: web.Url,
                        WebTitle: web.Title,
                        Id: userInfo.Id,
                        LoginName: userInfo.Name,
                        Name: userInfo.Title || userInfo.Name,
                        Email: userInfo.EMail,
                        Group: "",
                        GroupId: 0,
                        GroupInfo: "",
                        Role: role.RoleDefinitionBindings.results[0].Name,
                        RoleInfo: role.RoleDefinitionBindings.results[0].Description || ""
                    });

                    // Check the next role
                    continue;
                }
            }

            // Get the groups the user is associated with
            Web(DataSource.SiteContext.SiteFullUrl, { requestDigest: DataSource.SiteContext.FormDigestValue }).SiteUsers(userInfo.Id).Groups().execute(groups => {
                // Parse the groups the member belongs to
                Helper.Executor(groups.results, group => {
                    // Parse the roles
                    for (let i = 0; i < web.RoleAssignments.results.length; i++) {
                        let role: Types.SP.RoleAssignmentOData = web.RoleAssignments.results[i] as any;

                        // See if the user belongs to this role
                        if (role.Member.LoginName == group.LoginName) {
                            // Add the user information
                            this._items.push({
                                WebUrl: web.Url,
                                WebTitle: web.Title,
                                Id: userInfo.Id,
                                LoginName: userInfo.Name,
                                Name: userInfo.Title || userInfo.Name,
                                Email: userInfo.EMail,
                                Group: group.Title,
                                GroupId: group.Id,
                                GroupInfo: group.Description || "",
                                Role: role.RoleDefinitionBindings.results[0].Name,
                                RoleInfo: role.RoleDefinitionBindings.results[0].Description || ""
                            });
                        }
                    }
                }).then(resolve);
            }, resolve);
        });
    }

    // Gets the external users
    private static getUsers(): PromiseLike<IUserInfo[]> {
        // Return a promise
        return new Promise((resolve, reject) => {
            let users: IUserInfo[] = [];

            // Get the user information list
            Web(DataSource.SiteContext.SiteFullUrl, { requestDigest: DataSource.SiteContext.FormDigestValue }).Lists("User Information List").Items().query({
                Filter: `Title eq 'Everyone' or substringof('spo-grid-all-users', Name)`,
                Select: ["Id", "Name", "EMail", "Title", "UserName"],
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
                        Title: item.Title,
                        UserName: item["UserName"]
                    });
                }

                // Resolve the request
                resolve(users);
            }, reject);
        });
    }

    // Removes a user from a group
    private static removeUser(user: string, userId: number) {
        // Display a loading dialog
        LoadingDialog.setHeader("Removing User");
        LoadingDialog.setBody(`Removing the site user '${user}' from all groups. This will close after the request completes.`);
        LoadingDialog.show();

        // Remove the user from the site
        Web(DataSource.SiteContext.SiteFullUrl, { requestDigest: DataSource.SiteContext.FormDigestValue }).SiteUsers().removeById(userId).execute(
            // Success
            () => {
                // Close the dialog
                LoadingDialog.hide();
            },

            // Error
            () => {
                // Close the dialog
                LoadingDialog.hide();

                // TODO
            }
        )
    }

    // Removes a user from a group
    private static removeUserFromGroup(user: string, userId: number, groupId: number) {
        // Display a loading dialog
        LoadingDialog.setHeader("Removing User");
        LoadingDialog.setBody(`Removing the site user '${user}' from the group. This will close after the request completes.`);
        LoadingDialog.show();

        // Remove the user from the site
        Web(DataSource.SiteContext.SiteFullUrl, { requestDigest: DataSource.SiteContext.FormDigestValue }).SiteGroups().getById(groupId).Users().removeById(userId).execute(
            // Success
            () => {
                // Close the dialog
                LoadingDialog.hide();
            },

            // Error
            () => {
                // Close the dialog
                LoadingDialog.hide();

                // TODO
            }
        )
    }

    // Reverts the item permissions
    private static revertPermissions(item: ISearchItem) {
        // Show a loading dialog
        LoadingDialog.setHeader("Restoring Permissions");
        LoadingDialog.setBody("This window will close after the item permissions are restored...");
        LoadingDialog.show();

        // Restore the permissions
        Web(item.WebUrl, { requestDigest: DataSource.SiteContext.FormDigestValue })
            .Lists().getById(item.ListId).Items(item.ItemId).resetRoleInheritance().execute(() => {
                // Close the loading dialog
                LoadingDialog.hide();
            });
    }

    // Renders the search summary
    private static renderSummary(el: HTMLElement, auditOnly: boolean, items: ISearchItem[], onClose: () => void) {
        // Render the summary
        new Dashboard({
            el,
            navigation: {
                title: "Search EEEU",
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
                        new ExportCSV("searchEEEU.csv", CSVFields, items);
                    }
                }]
            },
            table: {
                rows: items,
                onRendering: dtProps => {
                    dtProps.columnDefs = [
                        {
                            "targets": 5,
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
                        name: "",
                        title: "Object Type",
                        onRenderCell: (el, col, item: ISearchItem) => {
                            // Set the object type
                            let objType = "Site";
                            if (item.GroupId > 0) { objType = "Group"; }
                            if (item.ItemId > 0) { objType = "Item"; }
                            if (item.FileUrl) { objType = "File" }

                            // Render the info
                            el.innerHTML = `
                                ${item.Name}
                                <br/>
                                <b>Type: </b>${objType}
                            `;
                        }
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
                        onRenderCell: (el, col, row: ISearchItem) => {
                            let btnDelete: Components.IButton = null;

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

                            // See if this is a list item
                            if (row.ListId) {
                                // Add the view button
                                tooltips.add({
                                    content: "Click to view the item unique permissions.",
                                    btnProps: {
                                        className: "pe-2 py-1",
                                        //iconType: GetIcon(24, 24, "PeopleTeam", "mx-1"),
                                        text: "View Group",
                                        type: Components.ButtonTypes.OutlinePrimary,
                                        onClick: () => {
                                            // View the group
                                            window.open(`${row.WebUrl}/${ContextInfo.layoutsUrl}/user.aspx?List=${row.ListId}&obj=${row.ListId},${row.FileUrl ? "1" : "2"},LISTITEM`);
                                        }
                                    }
                                });
                            }

                            // Ensure this is a group
                            if (row.GroupId > 0) {
                                // Add the view button
                                tooltips.add({
                                    content: "Click to view the site group.",
                                    btnProps: {
                                        className: "pe-2 py-1",
                                        //iconType: GetIcon(24, 24, "PeopleTeam", "mx-1"),
                                        text: "View Group",
                                        type: Components.ButtonTypes.OutlinePrimary,
                                        onClick: () => {
                                            // View the group
                                            window.open(`${row.WebUrl}/${ContextInfo.layoutsUrl}/people.aspx?MembershipGroupId=${row.GroupId}`);
                                        }
                                    }
                                });

                                // Ensure we can remove
                                if (!auditOnly) {
                                    // Add the remove button
                                    tooltips.add({
                                        content: "Click to remove the account from the group",
                                        btnProps: {
                                            assignTo: btn => { btnDelete = btn; },
                                            className: "pe-2 py-1",
                                            //iconType: GetIcon(24, 24, "PersonDelete", "mx-1"),
                                            text: "Remove From Group",
                                            type: Components.ButtonTypes.OutlineDanger,
                                            onClick: () => {
                                                // Confirm the removal of the user
                                                if (confirm("Are you sure you want to remove the account from this group?")) {
                                                    // Disable this button
                                                    btnDelete.disable();

                                                    // Remove the user
                                                    this.removeUserFromGroup(row.Name, row.Id, row.GroupId);
                                                }
                                            }
                                        }
                                    });
                                }
                            } else if (!auditOnly) {
                                // See if this is an item
                                if (row.ItemId > 0) {
                                    // Add the remove button
                                    tooltips.add({
                                        content: "Click to restore permissions to inherit from the parent",
                                        btnProps: {
                                            assignTo: btn => { btnDelete = btn; },
                                            className: "pe-2 py-1",
                                            //iconType: GetIcon(24, 24, "PeopleTeamDelete", "mx-1"),
                                            text: "Restore",
                                            type: Components.ButtonTypes.OutlineDanger,
                                            onClick: () => {
                                                // Confirm the deletion of the group
                                                if (confirm("Are you sure you restore the permissions to inherit?")) {
                                                    // Revert the permissions
                                                    this.revertPermissions(row);
                                                }
                                            }
                                        }
                                    });
                                } else {
                                    // Add the remove button
                                    tooltips.add({
                                        content: "Click to remove the account from all site groups and the site",
                                        btnProps: {
                                            assignTo: btn => { btnDelete = btn; },
                                            className: "pe-2 py-1",
                                            //iconType: GetIcon(24, 24, "PersonDelete", "mx-1"),
                                            text: "Remove From Site",
                                            type: Components.ButtonTypes.OutlineDanger,
                                            onClick: () => {
                                                // Confirm the removal of the user
                                                if (confirm("Are you sure you want to remove the account from this site?")) {
                                                    // Disable this button
                                                    btnDelete.disable();

                                                    // Remove the user
                                                    this.removeUser(row.Name, row.Id);
                                                }
                                            }
                                        }
                                    });
                                }
                            }
                        }
                    }
                ]
            }
        });
    }

    // Runs the report
    static run(el: HTMLElement, auditOnly: boolean, values: { [key: string]: any }, onClose: () => void) {
        // Show a loading dialog
        LoadingDialog.setHeader("Searching Site");
        LoadingDialog.setBody("Searching the current permissions of the site...");
        LoadingDialog.show();

        // Clear the items
        this._items = [];

        // See if we are showing hidden lists
        let searchLists = values["SearchLists"];

        // Parse all webs
        let counter = 0;
        Helper.Executor(DataSource.SiteItems, siteItem => {
            // Update the loading dialog
            LoadingDialog.setBody(`Getting the info for web ${++counter} of ${DataSource.SiteItems.length}...`);

            // Return a promise
            return new Promise(resolve => {
                // Get the permissions
                Web(siteItem.text, { requestDigest: DataSource.SiteContext.FormDigestValue }).query({
                    Expand: [
                        "RoleAssignments", "RoleAssignments/Groups", "RoleAssignments/Member",
                        "RoleAssignments/RoleDefinitionBindings", "SiteGroups"
                    ]
                }).execute(web => {
                    // Update the loading dialog
                    LoadingDialog.setBody(`Analyzing web ${counter} of ${DataSource.SiteItems.length}...`);

                    // Analyze the site
                    this.analyzeSite(web, searchLists).then(resolve);
                });
            });
        }).then(() => {
            // Clear the element
            while (el.firstChild) { el.removeChild(el.firstChild); }

            // Render the summary
            this.renderSummary(el, auditOnly, this._items, onClose);

            // Hide the loading dialog
            LoadingDialog.hide();
        });
    }
}