import { Dashboard, Documents, LoadingDialog } from "dattatable";
import { Components, ContextInfo, Helper, Search, Types, Web } from "gd-sprest-bs";
import { fileEarmark } from "gd-sprest-bs/build/icons/svgs/fileEarmark";
import { personX } from "gd-sprest-bs/build/icons/svgs/personX";
import { DataSource } from "../ds";
import { ExportCSV } from "./exportCSV";

interface IGroupInfo {
    DocUrl?: string;
    Role?: string;
    RoleInfo?: string;
    Name: string;
    Email?: string;
    Group?: string;
    GroupId?: number;
    GroupInfo?: string;
    Id?: number;
    WebTitle: string;
    WebUrl: string;
}

interface IUserInfo {
    EMail: string;
    Id: number;
    Name: string;
    Title: string;
}

const CSVFields = [
    "WebTitle", "WebUrl", "Name", "Email", "Group", "GroupInfo", "Role", "RoleInfo"
]

export class ExternalUsers {
    private static _items: IGroupInfo[] = [];

    // Analyzes the users
    private static analyzeUsers(rootWeb: Types.SP.WebOData, users: IUserInfo[]): PromiseLike<void> {
        // Return a promise
        return new Promise(resolve => {
            let counter = 0;

            // Parse the users
            Helper.Executor(users, user => {
                // Update the loading dialog
                LoadingDialog.setBody(`Analyzing User ${++counter} of ${users.length}`);

                // Get the user information
                return this.getUserInfo(rootWeb, user);
            }).then(() => {
                // Resolve the request
                resolve();
            });
        });
    }

    // Gets the form fields to display
    static getFormFields(): Components.IFormControlProps[] { return []; }

    // Get the user information
    private static getUserInfo(rootWeb: Types.SP.WebOData, userInfo: IUserInfo): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve) => {
            // Get the user information
            Web(DataSource.SiteContext.SiteFullUrl, { requestDigest: DataSource.SiteContext.FormDigestValue }).getUserById(userInfo.Id).query({
                Expand: ["Groups"]
            }).execute(user => {
                // Parse the groups
                Helper.Executor(user.Groups.results, group => {
                    // Return a promise
                    return new Promise((resolve) => {
                        // Parse the roles
                        for (let i = 0; i < rootWeb.RoleAssignments.results.length; i++) {
                            let role: Types.SP.RoleAssignmentOData = rootWeb.RoleAssignments.results[i] as any;

                            // See if the user belongs to this role
                            if (role.Member.LoginName == group.LoginName) {
                                // Add the user information
                                this._items.push({
                                    WebUrl: rootWeb.Url,
                                    WebTitle: rootWeb.Title,
                                    Id: userInfo.Id,
                                    Name: userInfo.Title || userInfo.Name,
                                    Email: userInfo.EMail,
                                    Group: group.Title,
                                    GroupId: group.Id,
                                    GroupInfo: group.Description || "",
                                    Role: role.RoleDefinitionBindings.results[0].Name,
                                    RoleInfo: role.RoleDefinitionBindings.results[0].Description || ""
                                });

                                // Resolve the request and return
                                resolve(null);
                                return;
                            }
                        }

                        // See if this is not a sharing link
                        if (!group.Title.startsWith("SharingLinks.")) {
                            // Add the user information
                            this._items.push({
                                WebUrl: rootWeb.Url,
                                WebTitle: rootWeb.Title,
                                Name: userInfo.Title || userInfo.Name,
                                Email: userInfo.EMail,
                                Group: group.Title,
                                GroupId: group.Id,
                                GroupInfo: group.Description || "",
                                Role: "Unknown",
                                RoleInfo: "Unable to determine the role for this group"
                            });

                            // Resolve the promise
                            resolve(null);
                            return;
                        }

                        // Set the role name
                        let info = group.Title.split('.');
                        let docId = info.length > 2 ? info[1] : "";
                        let roleName = info.length > 3 ? info[2] : "";

                        // See if the doc id exists
                        if (docId) {
                            // Find the document by its id
                            Search.postQuery<{ Path: string; SPWebUrl: string; Title: string }>({
                                url: DataSource.SiteContext.SiteFullUrl,
                                query: {
                                    Querytext: "UniqueID: " + docId,
                                    SelectProperties: {
                                        results: ["Path", "SPWebUrl", "Title"]
                                    }
                                },
                                targetInfo: { requestDigest: DataSource.SiteContext.FormDigestValue }
                            }).then(search => {
                                let result = search.results[0];

                                // Set the role information
                                let roleInfo = "";
                                if (result) {
                                    // Set the role information
                                    roleInfo = "Has '" + roleName + "' access to the file <a target='_blank' " +
                                        "href='" + result.Path + "'>" + result.Title + "</a>.";
                                }

                                // Add the user information
                                this._items.push({
                                    DocUrl: result?.Path,
                                    WebUrl: result?.SPWebUrl,
                                    WebTitle: "",
                                    Name: userInfo.Title || userInfo.Name,
                                    Email: userInfo.EMail,
                                    Group: group.Title,
                                    GroupId: group.Id,
                                    GroupInfo: group.Description,
                                    Role: roleName,
                                    RoleInfo: roleInfo
                                });
                            });
                        }
                    });
                }).then(() => {
                    // Resolve the promise
                    resolve();
                });
            });
        });
    }

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

    // Removes a user from a group
    private static removeUser(user: string, userId: number, group: string) {
        // Display a loading dialog
        LoadingDialog.setHeader("Removing Site User");
        LoadingDialog.setBody(`Removing the site user '${user}' from the '${group}' group. This will close after the request completes.`);
        LoadingDialog.show();

        // Remove the user from the group
        Web(DataSource.SiteContext.SiteFullUrl, { requestDigest: DataSource.SiteContext.FormDigestValue })
            .SiteGroups(group).Users().removeById(userId).execute(
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
                title: "Search Content",
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
                            "targets": 7,
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
                        name: "Name",
                        title: "User Name"
                    },
                    {
                        name: "Email",
                        title: "User Email"
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

                            // Ensure this is a group
                            if (row.Group) {
                                // Render the buttons
                                Components.TooltipGroup({
                                    el,
                                    tooltips: [
                                        {
                                            content: "View Group",
                                            btnProps: {
                                                className: "pe-2 py-1",
                                                iconType: fileEarmark(24, 24, "mx-1"),
                                                text: "View",
                                                type: Components.ButtonTypes.OutlinePrimary,
                                                isDisabled: !(row.GroupId > 0),
                                                onClick: () => {
                                                    // View the group
                                                    window.open(`${row.WebUrl}/${ContextInfo.layoutsUrl}/people.aspx?MembershipGroupId=${row.GroupId}`);
                                                }
                                            }
                                        },
                                        {
                                            content: "Remove User",
                                            btnProps: {
                                                assignTo: btn => { btnDelete = btn; },
                                                className: "pe-2 py-1",
                                                iconType: personX(24, 24, "mx-1"),
                                                text: "Remove",
                                                type: Components.ButtonTypes.OutlineDanger,
                                                isDisabled: !(row.Id > 0),
                                                onClick: () => {
                                                    // Confirm the deletion of the group
                                                    if (confirm("Are you sure you want to remove the user from this group?")) {
                                                        // Disable this button
                                                        btnDelete.disable();

                                                        // Delete the site group
                                                        this.removeUser(row.Name, row.Id, row.Group);
                                                    }
                                                }
                                            }
                                        }
                                    ]
                                });
                            } else {
                                // Render the delete button
                                Components.Tooltip({
                                    el,
                                    content: "Delete Group",
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
                            }
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
            LoadingDialog.setBody("Loading the site users...");

            // Get the users for this site
            Web(DataSource.SiteContext.SiteFullUrl, { requestDigest: DataSource.SiteContext.FormDigestValue }).Lists("User Information List").Items().query({
                Filter: "substringof('%23ext%23', Name)",
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
                LoadingDialog.setHeader("Analyzing Users/Groups");

                // Analyze the users
                this.analyzeUsers(web, users).then(() => {
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