import { Dashboard, LoadingDialog } from "dattatable";
import { Components, ContextInfo, Helper, Types, Web, SPTypes } from "gd-sprest-bs";
import { DataSource } from "../ds";
import { ExportCSV } from "./exportCSV";

interface ISearchItem {
    Role?: string;
    RoleInfo?: string;
    LoginName: string;
    Name: string;
    Email?: string;
    Group?: string;
    GroupId?: number;
    GroupInfo?: string;
    Id?: number;
    WebTitle: string;
    WebUrl: string;
    ListName?: string;
    ListUrl?: string;
    ItemId?: number;
    ItemType?: string;
}

interface IUserInfo {
    EMail: string;
    Id: number;
    Name: string;
    Title: string;
    UserName: string;
}

const CSVFields = [
    "Name", "UserName", "Email", "Group", "GroupInfo",
    "Role", "RoleInfo", "WebTitle", "WebUrl", "ListName", "ListUrl", "ItemId", "ItemType"
]

export class SearchEEEU {
    private static _items: ISearchItem[] = null;

    // Analyzes the lists and libraries with unique permissions
    private static analyzeListsAndLibraries(web: Types.SP.WebOData, users: IUserInfo[]) {
        // Return a promise
        return new Promise(resolve => {
            // Get the lists with unique permissions for this web
            Web(web.Url, { requestDigest: DataSource.SiteContext.FormDigestValue }).Lists().query({
                Filter: "Hidden eq false and HasUniqueRoleAssignments eq true",
                Expand: ["RoleAssignments", "RoleAssignments/Member", "RoleAssignments/RoleDefinitionBindings"],
                Select: ["Title", "RootFolder/ServerRelativeUrl", "HasUniqueRoleAssignments"]
            }).execute(lists => {
                let counter = 0;

                // Parse the lists with unique permissions
                Helper.Executor(lists.results, list => {
                    // Update the loading dialog
                    LoadingDialog.setBody(`Analyzing List ${++counter} of ${lists.results.length} in web: ${web.Title}`);

                    // Check list-level permissions
                    return this.checkListPermissions(web, list, users);
                }).then(resolve);
            }, resolve);
        });
    }

    // Checks permissions on a specific list
    private static checkListPermissions(web: Types.SP.WebOData, list: Types.SP.ListOData, users: IUserInfo[]) {
        // Return a promise
        return new Promise(resolve => {
            // Parse the users to check against list permissions
            Helper.Executor(users, user => {
                // Return a promise
                return new Promise(resolve => {
                    // Parse the role assignments for this list
                    for (let i = 0; i < list.RoleAssignments.results.length; i++) {
                        let role: Types.SP.RoleAssignmentOData = list.RoleAssignments.results[i] as any;

                        // See if this role is the user directly
                        if (role.Member.LoginName == user.Name) {
                            // Add the user information
                            this._items.push({
                                WebUrl: web.Url,
                                WebTitle: web.Title,
                                Id: user.Id,
                                LoginName: user.Name,
                                Name: user.Title || user.Name,
                                Email: user.EMail,
                                Group: "",
                                GroupId: 0,
                                GroupInfo: "",
                                Role: role.RoleDefinitionBindings.results[0].Name,
                                RoleInfo: role.RoleDefinitionBindings.results[0].Description || "",
                                ListName: list.Title,
                                ListUrl: list.RootFolder?.ServerRelativeUrl,
                                ItemType: "List"
                            });
                        }
                    }

                    // Get the groups the user is associated with
                    Web(DataSource.SiteContext.SiteFullUrl, { requestDigest: DataSource.SiteContext.FormDigestValue }).SiteUsers(user.Id).Groups().execute(groups => {
                        // Parse the groups the member belongs to
                        Helper.Executor(groups.results, group => {
                            // Parse the role assignments for this list
                            for (let i = 0; i < list.RoleAssignments.results.length; i++) {
                                let role: Types.SP.RoleAssignmentOData = list.RoleAssignments.results[i] as any;

                                // See if the user belongs to this role through group membership
                                if (role.Member.LoginName == group.LoginName) {
                                    // Add the user information
                                    this._items.push({
                                        WebUrl: web.Url,
                                        WebTitle: web.Title,
                                        Id: user.Id,
                                        LoginName: user.Name,
                                        Name: user.Title || user.Name,
                                        Email: user.EMail,
                                        Group: group.Title,
                                        GroupId: group.Id,
                                        GroupInfo: group.Description || "",
                                        Role: role.RoleDefinitionBindings.results[0].Name,
                                        RoleInfo: role.RoleDefinitionBindings.results[0].Description || "",
                                        ListName: list.Title,
                                        ListUrl: list.RootFolder?.ServerRelativeUrl,
                                        ItemType: "List"
                                    });
                                }
                            }
                        }).then(resolve);
                    }, resolve);
                });
            }).then(resolve);
        });
    }

    // Analyzes the site
    private static analyzeSite(web: Types.SP.WebOData) {
        // Return a promise
        return new Promise(resolve => {
            // Show a loading dialog
            LoadingDialog.setBody("Getting the user information...");

            // Get the users
            this.getUsers().then(users => {
                let counter = 0;

                // Parse the users
                Helper.Executor(users, user => {
                    // Update the loading dialog
                    LoadingDialog.setBody(`Analyzing User ${++counter} of ${users.length} for web: ${web.Title}`);

                    // Get the user information
                    return this.getUserInfo(web, user);
                }).then(() => {
                    // Resolve the request
                    resolve(null);
                });
            }, resolve);
        });
    }

    // Gets the form fields to display
    static getFormFields(): Components.IFormControlProps[] { 
        return [
            {
                name: "RecursiveSearch",
                label: "Recursive Search",
                description: "Enable this option to search all sub-webs in the site collection, including lists and libraries with unique permissions.",
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
                        RoleInfo: role.RoleDefinitionBindings.results[0].Description || "",
                        ListName: "",
                        ListUrl: "",
                        ItemType: "Web"
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
                                RoleInfo: role.RoleDefinitionBindings.results[0].Description || "",
                                ListName: "",
                                ListUrl: "",
                                ItemType: "Web"
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
    // Renders the search summary
    private static renderSummary(el: HTMLElement, auditOnly: boolean, items: ISearchItem[], onClose: () => void) {
        // Render the summary
        new Dashboard({
            el,
            navigation: {
                title: "Search Users",
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
                        new ExportCSV("searchDocs.csv", CSVFields, items);
                    }
                }]
            },
            table: {
                rows: items,
                onRendering: dtProps => {
                    dtProps.columnDefs = [
                        {
                            "targets": 7,
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
                        name: "Name",
                        title: "User Name"
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
                        name: "ItemType",
                        title: "Item Type"
                    },
                    {
                        name: "ListName",
                        title: "List/Library"
                    },
                    {
                        className: "text-end",
                        name: "",
                        title: "",
                        onRenderCell: (el, col, row: ISearchItem) => {
                            let btnDelete: Components.IButton = null;

                            // Render the tooltips
                            let tooltips = Components.TooltipGroup({ el });

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
                                        isDisabled: !(row.GroupId > 0),
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

        // Check if recursive search is enabled
        let recursiveSearch = values["RecursiveSearch"];

        if (recursiveSearch) {
            // Recursively search all webs in the site collection
            this.runRecursiveSearch().then(() => {
                // Clear the element
                while (el.firstChild) { el.removeChild(el.firstChild); }

                // Render the summary
                this.renderSummary(el, auditOnly, this._items, onClose);

                // Hide the loading dialog
                LoadingDialog.hide();
            });
        } else {
            // Run the search only on the current web
            this.runSingleWebSearch().then(() => {
                // Clear the element
                while (el.firstChild) { el.removeChild(el.firstChild); }

                // Render the summary
                this.renderSummary(el, auditOnly, this._items, onClose);

                // Hide the loading dialog
                LoadingDialog.hide();
            });
        }
    }

    // Runs the search on a single web (current behavior)
    private static runSingleWebSearch(): PromiseLike<void> {
        // Return a promise
        return new Promise(resolve => {
            // Get the permissions for the current web
            Web(DataSource.SiteContext.SiteFullUrl, { requestDigest: DataSource.SiteContext.FormDigestValue }).query({
                Expand: [
                    "RoleAssignments", "RoleAssignments/Groups", "RoleAssignments/Member",
                    "RoleAssignments/Member/Users", "RoleAssignments/RoleDefinitionBindings",
                    "SiteGroups"
                ]
            }).execute(web => {
                // Analyze the site
                this.analyzeSite(web).then(() => {
                    // Also check lists and libraries with unique permissions in this web
                    this.getUsers().then(users => {
                        this.analyzeListsAndLibraries(web, users).then(resolve);
                    });
                });
            });
        });
    }

    // Runs the recursive search across all webs in the site collection
    private static runRecursiveSearch(): PromiseLike<void> {
        // Return a promise
        return new Promise(resolve => {
            // Parse all the webs in the site collection
            let counter = 0;
            Helper.Executor(DataSource.SiteItems, siteItem => {
                // Update the loading dialog
                LoadingDialog.setHeader(`Searching Web ${++counter} of ${DataSource.SiteItems.length}`);
                LoadingDialog.setBody(`Analyzing permissions for: ${siteItem.text}`);

                // Return a promise
                return new Promise(resolve => {
                    // Get the permissions for this web
                    Web(siteItem.text, { requestDigest: DataSource.SiteContext.FormDigestValue }).query({
                        Expand: [
                            "RoleAssignments", "RoleAssignments/Groups", "RoleAssignments/Member",
                            "RoleAssignments/Member/Users", "RoleAssignments/RoleDefinitionBindings",
                            "SiteGroups"
                        ]
                    }).execute(web => {
                        // Analyze the site-level permissions
                        this.analyzeSite(web).then(() => {
                            // Get the users to check against list permissions
                            this.getUsers().then(users => {
                                // Also check lists and libraries with unique permissions
                                this.analyzeListsAndLibraries(web, users).then(resolve);
                            });
                        });
                    }, resolve);
                });
            }).then(resolve);
        });
    }
}