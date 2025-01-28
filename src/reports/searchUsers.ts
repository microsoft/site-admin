import { Dashboard, LoadingDialog } from "dattatable";
import { Components, ContextInfo, Helper, Types, Web } from "gd-sprest-bs";
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
    "Role", "RoleInfo", "WebTitle", "WebUrl"
]

export class SearchUsers {
    private static _items: ISearchItem[] = null;

    // Analyzes the site
    private static analyzeSite(web: Types.SP.WebOData, user: string | Types.SP.User) {
        // Return a promise
        return new Promise(resolve => {
            // Show a loading dialog
            LoadingDialog.setBody("Getting the user information...");

            // Get the users
            this.getUsers(user).then(users => {
                let counter = 0;

                // Parse the users
                Helper.Executor(users, user => {
                    // Update the loading dialog
                    LoadingDialog.setBody(`Analyzing User ${++counter} of ${users.length}`);

                    // Get the user information
                    return this.getUserInfo(web, user);
                }).then(() => {
                    // Hide the loading dialog
                    LoadingDialog.hide();

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
                label: "User or Group Search by Text",
                name: "UserName",
                className: "mb-3",
                description: "Type a user or group display or login name for the search",
                errorMessage: "Please enter the user or group information",
                type: Components.FormControlTypes.TextField
            },
            {
                label: "People Search by Lookup",
                name: "PeoplePicker",
                className: "mb-3",
                description: "Enter a minimum of 3 characters to search for a user",
                errorMessage: "No user was selected...",
                allowGroups: false,
                type: Components.FormControlTypes.PeoplePicker
            } as Components.IFormControlPropsPeoplePicker
        ];
    }

    // Get the user information
    private static getUserInfo(web: Types.SP.WebOData, userInfo: IUserInfo) {
        // Return a promise
        return new Promise((resolve) => {
            // Parse the groups
            Helper.Executor(web.SiteGroups.results, group => {
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
            }).then(() => {
                // Resolve the promise
                resolve(null);
            });
        });
    }

    // Gets the external users
    private static getUsers(user: string | Types.SP.User): PromiseLike<IUserInfo[]> {
        // Return a promise
        return new Promise((resolve, reject) => {
            let users: IUserInfo[] = [];

            // See if we are searching by a string
            if (typeof (user) === "string") {
                // Get the user information list
                Web(DataSource.SiteContext.SiteFullUrl, { requestDigest: DataSource.SiteContext.FormDigestValue }).Lists("User Information List").Items().query({
                    Filter: `substringof('${user}', Name) or substringof('${user}', Title) or substringof('${user}', UserName)`,
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
            } else {
                // Get the users
                // TODO - See if we can use the filter option
                Web(DataSource.SiteContext.SiteFullUrl, { requestDigest: DataSource.SiteContext.FormDigestValue }).Lists("User Information List").Items().query({
                    GetAllItems: true,
                    Select: ["Id", "Name", "EMail", "Title", "UserName"],
                    Top: 5000
                }).execute(items => {
                    // Parse the items
                    for (let i = 0; i < items.results.length; i++) {
                        let item = items.results[i];

                        // See if this is the target user
                        if (item["EMail"] == user.Email || item["Name"] == user.LoginName || item["UserName"] == user.UserPrincipalName) {
                            // Add the user
                            users.push({
                                EMail: item["EMail"],
                                Id: item.Id,
                                Name: item["Name"],
                                Title: item.Title,
                                UserName: item["UserName"]
                            });
                        }
                    }

                    // Resolve the request
                    resolve(users);
                }, reject);
            }
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

    // Renders the search summary
    private static renderSummary(el: HTMLElement, items: ISearchItem[], onClose: () => void) {
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
                        name: "LoginName",
                        title: "Login"
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
                        onRenderCell: (el, col, row: ISearchItem) => {
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
                                                //iconType: GetIcon(24, 24, "PeopleTeam", "mx-1"),
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
                                                //iconType: GetIcon(24, 24, "PersonDelete", "mx-1"),
                                                text: "Remove",
                                                type: Components.ButtonTypes.OutlineDanger,
                                                isDisabled: !(row.Id > 0),
                                                onClick: () => {
                                                    // Confirm the removal of the user
                                                    if (confirm("Are you sure you want to remove the user from this site?")) {
                                                        // Disable this button
                                                        btnDelete.disable();

                                                        // Remove the user
                                                        this.removeUser(row.Name, row.Id);
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
                                    content: "Remove User",
                                    btnProps: {
                                        assignTo: btn => { btnDelete = btn; },
                                        className: "pe-2 py-1",
                                        //iconType: GetIcon(24, 24, "PersonDelete", "mx-1"),
                                        text: "Remove",
                                        type: Components.ButtonTypes.OutlineDanger,
                                        onClick: () => {
                                            // Confirm the deletion of the group
                                            if (confirm("Are you sure you want to remove the user from this site?")) {
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
    static run(el: HTMLElement, values: { [key: string]: any }, onClose: () => void) {
        // Show a loading dialog
        LoadingDialog.setHeader("Searching Site");
        LoadingDialog.setBody("Searching the current permissions of the site...");
        LoadingDialog.show();

        // Clear the items
        this._items = [];

        // Get the form values
        let userName: string = values["UserName"].trim();
        let user: Types.SP.User = values["PeoplePicker"][0];

        // Get the permissions
        Web(DataSource.SiteContext.SiteFullUrl, { requestDigest: DataSource.SiteContext.FormDigestValue }).query({
            Expand: [
                "RoleAssignments", "RoleAssignments/Groups", "RoleAssignments/Member",
                "RoleAssignments/Member/Users", "RoleAssignments/RoleDefinitionBindings"
            ]
        }).execute(web => {
            // Analyze the site
            this.analyzeSite(web, userName || user).then(() => {
                // Clear the element
                while (el.firstChild) { el.removeChild(el.firstChild); }

                // Render the summary
                this.renderSummary(el, this._items, onClose);

                // Hide the loading dialog
                LoadingDialog.hide();
            });
        });
    }
}