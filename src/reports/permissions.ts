import { Dashboard, LoadingDialog, Modal } from "dattatable";
import { Components, ContextInfo, DirectorySession, Helper, SPTypes, Types, Web } from "gd-sprest-bs";
import { DataSource } from "../ds";
import { ExportCSV } from "./exportCSV";

interface IPermissionItem {
    EEEU?: boolean;
    Everyone?: boolean;
    GroupIds?: string[];
    GroupMembers?: Types.SP.Directory.User[];
    GroupMembersAsString?: string[];
    GroupOwners?: Types.SP.Directory.User[];
    GroupOwnersAsString?: string[];
    GroupUrl?: string;
    Id?: number;
    LoginName: string;
    Name: string;
    Roles?: string[];
    RoleInfo?: string[];
    SiteMembers?: Types.SP.User[];
    SiteMembersAsString?: string[];
    SiteOwners?: Types.SP.User[];
    SiteOwnersAsString?: string[];
    Type: "AD Group" | "M365 Group" | "Site Group" | "User";
    WebTitle: string;
    WebUrl: string;
}

const CSVFields = [
    "WebUrl", "WebTitle", "Name", "Type", "Id", "LoginName", "Everyone", "EEEU", "GroupIds", "Roles",
    "GroupMembersAsString", "GroupOwnersAsString", "SiteMembersAsString", "SiteOwnersAsString"
]

export class Permissions {
    private static _groupIds: { [key: string]: Types.SP.Directory.GroupOData } = null;
    private static _groupIdErrors: string[] = null;
    private static _roleMapper: { [key: string]: Types.SP.User } = null;
    private static _items: IPermissionItem[] = null;

    // Analyze the group ids
    private static analyzeGroupIds(groupIds: string[]): PromiseLike<void> {
        // Return a promise
        return new Promise(resolve => {
            // Try to get the groups
            let counter = 0;
            Helper.Executor(groupIds, groupId => {
                // Update the loading dialog
                LoadingDialog.setBody(`Trying to get M365 Group information ${++counter} of ${groupIds.length}`);

                // Return a promise
                return new Promise(resolve => {
                    // Skip this, if we have already queried for this group
                    if (this._groupIds[groupId] || this._groupIdErrors.indexOf(groupId) >= 0) { resolve(null); return; }

                    // Get the group information
                    DirectorySession().group(groupId).query({
                        Expand: ["members", "owners"],
                        Select: [
                            "calendarUrl", "displayName", "id", "isPublic", "mail",
                            "members/principalName", "members/id", "members/displayName", "members/mail",
                            "owners/principalName", "owners/id", "owners/displayName", "owners/mail"
                        ]
                    }).execute(group => {
                        // Add the group information to the mapper
                        this._groupIds[group.id] = group;

                        // Try the next group
                        resolve(null);
                    }, () => {
                        // Append the error
                        this._groupIdErrors.push(groupId);

                        // Try the next group
                        resolve(null);
                    });
                });
            }).then(() => {
                // Update the loading dialog
                LoadingDialog.setBody(`Updating the role assignments with the M365 Group information`);

                // Parse the items
                this._items.forEach(item => {
                    // Clear the groups
                    item.GroupMembers = [];
                    item.GroupMembersAsString = [];
                    item.GroupOwners = [];
                    item.GroupOwnersAsString = [];
                    item.SiteMembersAsString = [];
                    item.SiteOwnersAsString = [];

                    // Ensure groups exist
                    if (item.GroupIds.length == 0) { return; }

                    // Parse the group ids
                    item.GroupIds.forEach(groupId => {
                        // Get the group
                        let group = this._groupIds[groupId];
                        if (group) {
                            // Add the members/owners
                            item.GroupMembers = item.GroupMembers.concat(group.members.results);
                            item.GroupOwners = item.GroupOwners.concat(group.owners.results);
                            item.GroupUrl = group.calendarUrl.replace("calendar/group", "groups") + "/members";
                        }
                    });

                    // Parse the group members
                    item.GroupMembers.forEach(member => {
                        // Add the member name
                        item.GroupMembersAsString.push(member.mail);
                    });

                    // Remove duplicates
                    item.GroupMembersAsString = item.GroupMembersAsString.filter((value, idx, self) => {
                        return self.indexOf(value) === idx;
                    });

                    // Parse the group owners
                    item.GroupOwners.forEach(owner => {
                        // Add the member name
                        item.GroupOwnersAsString.push(owner.mail);
                    });

                    // Remove duplicates
                    item.GroupOwnersAsString = item.GroupOwnersAsString.filter((value, idx, self) => {
                        return self.indexOf(value) === idx;
                    });

                    // Parse the site members
                    item.SiteMembers.forEach(member => {
                        // Add the member name
                        item.SiteMembersAsString.push(member.Email || member.UserPrincipalName || member.LoginName);
                    });

                    // Remove duplicates
                    item.SiteMembersAsString = item.SiteMembersAsString.filter((value, idx, self) => {
                        return self.indexOf(value) === idx;
                    });

                    // Parse the site owners
                    item.SiteOwners.forEach(owner => {
                        // Add the member name
                        item.SiteOwnersAsString.push(owner.Email || owner.UserPrincipalName || owner.LoginName);
                    });

                    // Remove duplicates
                    item.SiteOwnersAsString = item.SiteOwnersAsString.filter((value, idx, self) => {
                        return self.indexOf(value) === idx;
                    });
                });

                // Resolve the request
                resolve();
            });
        });
    }

    // Analyze the role
    private static analyzeRole(role: Types.SP.RoleAssignmentOData, webUrl: string, webTitle: string): Promise<string[]> {
        // Expands the owners and users
        let getOwnerAndUsersForRole = () => {
            return new Promise(resolve => {
                // Do not do this for user types
                if (role.Member.PrincipalType == SPTypes.PrincipalTypes.User) { resolve(null); return; }

                // Get the owner information
                Web(webUrl).RoleAssignments(role.PrincipalId).query({ Expand: ["Member/Owner", "Member/Users"], Select: ["Member/Owner", "Member/Users"] }).execute(roleQuery => {
                    // Update the owner and users
                    role.Member["Owner"] = roleQuery.Member["Owner"];
                    role.Member["Users"] = roleQuery.Member["Users"];

                    // Resolve the request
                    resolve(null);
                }, resolve);
            });
        }

        // Return a promise
        return new Promise(resolve => {
            let item: IPermissionItem = null;

            // Get the owner and users for the role
            getOwnerAndUsersForRole().then(() => {
                // See if this is a user
                if (role.Member.PrincipalType == SPTypes.PrincipalTypes.User) {
                    // Add the role information
                    item = {
                        EEEU: false,
                        Everyone: false,
                        GroupIds: [],
                        Id: role.Member.Id,
                        LoginName: role.Member.LoginName,
                        Name: role.Member.Title,
                        Roles: [],
                        RoleInfo: [],
                        SiteMembers: [],
                        SiteOwners: [],
                        Type: "User",
                        WebTitle: webTitle,
                        WebUrl: webUrl
                    };
                }
                // Else, see if this is a site group
                else if (role.Member.PrincipalType == SPTypes.PrincipalTypes.SharePointGroup) {
                    // Add the role information
                    item = {
                        EEEU: false,
                        Everyone: false,
                        GroupIds: [],
                        Id: role.Member.Id,
                        LoginName: role.Member.LoginName,
                        Name: role.Member.Title,
                        Roles: [],
                        RoleInfo: [],
                        SiteMembers: role.Member["Users"] ? role.Member["Users"].results : [],
                        SiteOwners: role.Member["Owner"] ? [role.Member["Owner"]] : [],
                        Type: "Site Group",
                        WebTitle: webTitle,
                        WebUrl: webUrl,
                    };
                } else {
                    // See if this is a M365 group
                    let groupId = DataSource.getGroupId(role.Member.LoginName);

                    // Add the role information
                    item = {
                        EEEU: role.Member.Title == "Everyone except external users",
                        Everyone: role.Member.Title == "Everyone",
                        GroupIds: groupId ? [groupId] : [],
                        Id: role.Member.Id,
                        LoginName: role.Member.LoginName,
                        Name: role.Member.Title,
                        Roles: [],
                        RoleInfo: [],
                        SiteMembers: role.Member["Users"] ? role.Member["Users"].results : [],
                        SiteOwners: role.Member["Owner"] ? [role.Member["Owner"]] : [],
                        Type: groupId ? "M365 Group" : "AD Group",
                        WebTitle: webTitle,
                        WebUrl: webUrl
                    };
                }

                // Parse the role definitions
                role.RoleDefinitionBindings.results.forEach(roleDef => {
                    // Add the permission
                    item.Roles.push(roleDef.Name);
                    item.RoleInfo.push(roleDef.Description);
                });

                // Analyze the users for this item
                this.analyzeUsers(item);

                // Reolve the group ids
                resolve(item.GroupIds);
            });
        });
    }

    // Analyzes the roles
    private static analyzeRoles(roles: Types.SP.RoleAssignmentOData[], webUrl: string, webTitle: string): PromiseLike<void> {
        // Return a promise
        return new Promise(resolve => {
            let counter = 0;
            let groupIds: string[] = [];

            // Show a loading dialog
            LoadingDialog.setBody("Getting the user information...");

            // Parse the roles
            Helper.Executor(roles, role => {
                // Update the loading dialog
                LoadingDialog.setBody(`Analyzing Role ${++counter} of ${roles.length}`);

                // Return a promise
                return new Promise(resolve => {
                    // Analyze the role
                    this.analyzeRole(role, webUrl, webTitle).then(ids => {
                        // Add the group ids
                        groupIds = groupIds.concat(ids);

                        // Resolve the request
                        resolve(null);
                    });
                });
            }).then(() => {
                // Remove duplicates from the array
                groupIds = groupIds.filter((value, idx, self) => {
                    return self.indexOf(value) === idx;
                });

                // Analyze the group ids
                this.analyzeGroupIds(groupIds).then(() => {
                    // Hide the loading dialog
                    LoadingDialog.hide();

                    // Resolve the values
                    resolve();
                });
            });
        });
    }

    // Analyze the users for this item
    private static analyzeUsers(item: IPermissionItem) {
        // Return if the item doesn't exist
        if (item == null) { return; }

        // Add the item
        this._items.push(item);

        // Parse the users
        let users: Types.SP.User[] = (item.SiteMembers || []).concat(item.SiteOwners || []);
        for (let i = 0; i < users.length; i++) {
            let user = users[i];

            // Set the everyone flags
            if (user.Title == "Everyone") {
                // Set the flag
                item.Everyone = true;
            }
            if (user.Title == "Everyone except external users") {
                // Set the flag
                item.EEEU = true;
            }

            // See if this is a M365 Group
            if (user.PrincipalType == SPTypes.PrincipalTypes.SecurityGroup) {
                // Get the group id
                let groupId = DataSource.getGroupId(user.LoginName);
                if (groupId) {
                    // Add the group id
                    item.GroupIds.push(groupId);
                    this._roleMapper[groupId] = user;
                }
            }
        }
    }

    // Gets the form fields to display
    static getFormFields(): Components.IFormControlProps[] { return []; }

    // Renders the errors
    private static renderErrors() {
        // Clear the modal
        Modal.clear();

        // Set the header
        Modal.setHeader("Permission Errors");

        // Parse the errors
        let rows = [];
        for (let i = 0; i < this._groupIdErrors.length; i++) {
            let groupId = this._groupIdErrors[i];
            let mapper = this._roleMapper[groupId];

            // Add the row
            rows.push({
                id: groupId,
                title: mapper ? mapper.Title : "",
            });
        }

        // Render a dashboard
        new Dashboard({
            el: Modal.BodyElement,
            navigation: {
                title: "M365 Group Errors",
                showFilter: false,
                showSearch: false
            },
            table: {
                rows,
                columns: [
                    {
                        name: "id",
                        title: "Group Id"
                    },
                    {
                        name: "title",
                        title: "Group Name"
                    },
                    {
                        name: "",
                        onRenderCell: (el) => {
                            el.innerHTML = "Unable to get the group information.";
                        }
                    }
                ]
            }
        });

        // Render the footer
        Components.Button({
            text: "Close",
            onClick: () => {
                Modal.hide();
            }
        });

        // Show the modal
        Modal.show();
    }

    // Renders the search summary
    private static renderSummary(el: HTMLElement, auditOnly: boolean, items: IPermissionItem[], onClose: () => void) {
        // Create the nav items
        let navItems: Components.INavbarItem[] = [{
            text: "New Search",
            className: "btn-outline-light",
            isButton: true,
            onClick: () => {
                // Call the close event
                onClose();
            }
        }];

        // See if there are errors
        if (this._groupIdErrors.length > 0) {
            // Show the error button
            navItems.push({
                text: "Errors",
                className: "btn-outline-light ms-2",
                isButton: true,
                onClick: () => {
                    // Display the errors
                    this.renderErrors();
                }
            })
        }

        // Render the summary
        new Dashboard({
            el,
            navigation: {
                title: "Permissions",
                showFilter: false,
                items: navItems,
                itemsEnd: [{
                    text: "Export to CSV",
                    className: "btn-outline-light me-2",
                    isButton: true,
                    onClick: () => {
                        // Export the CSV
                        new ExportCSV("permissions.csv", CSVFields, items);
                    }
                }]
            },
            table: {
                rows: items,
                onRendering: dtProps => {
                    dtProps.columnDefs = [
                        {
                            "targets": 8,
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
                        title: "Role Name",
                        onRenderCell: (el, col, item: IPermissionItem) => {
                            // Render a view link
                            Components.Tooltip({
                                el,
                                content: "Click to view the group/user information in another tab.",
                                btnProps: {
                                    text: item.Name,
                                    type: Components.ButtonTypes.OutlinePrimary,
                                    onClick: () => {
                                        let url: string = null;

                                        // Display the users
                                        switch (item.Type) {
                                            case "AD Group":
                                                url = `${item.WebUrl}/${ContextInfo.layoutsUrl}/userdisp.aspx?ID=${item.Id}`;
                                                break;
                                            case "M365 Group":
                                                url = item.GroupUrl;
                                                break;
                                            case "Site Group":
                                                url = `${item.WebUrl}/${ContextInfo.layoutsUrl}/people.aspx?MembershipGroupId=${item.Id}`;
                                                break;
                                            case "User":
                                                url = `${item.WebUrl}/${ContextInfo.layoutsUrl}/userdisp.aspx?ID=${item.Id}`;
                                                break;
                                        }

                                        // Open the url in a new tab
                                        window.open(url, "_blank");
                                    }
                                }
                            });
                        }
                    },
                    {
                        name: "Type",
                        title: "Role Type"
                    },
                    {
                        name: "Everyone",
                        title: "Has Everyone?"
                    },
                    {
                        name: "EEEU",
                        title: "Has EEEU?"
                    },
                    {
                        name: "",
                        title: "M365 Groups",
                        onRenderCell: (el, col, item: IPermissionItem) => {
                            // Set the value
                            el.innerHTML = item.GroupIds.length.toString();

                            // Set the filter and order values
                            el.setAttribute("data-filter", el.innerHTML);
                            el.setAttribute("data-order", el.innerHTML);
                        }
                    },
                    {
                        name: "",
                        title: "M365 Users",
                        onRenderCell: (el, col, item: IPermissionItem) => {
                            // Set the value
                            el.innerHTML = item.GroupMembersAsString.length.toString();

                            // Set the filter and order values
                            el.setAttribute("data-filter", el.innerHTML);
                            el.setAttribute("data-order", el.innerHTML);
                        }
                    },
                    {
                        name: "",
                        title: "AD Accounts",
                        onRenderCell: (el, col, item: IPermissionItem) => {
                            // Ensure the site members exist
                            if (item.SiteMembers) {
                                // Get the non user accounts
                                let users = item.SiteMembers.filter(value => {
                                    // See if this is a user
                                    if (value.PrincipalType == SPTypes.PrincipalTypes.User) {
                                        // Determine if this is an AD account by the email
                                        return value.Email ? false : true;
                                    }

                                    // Exclude M365 groups
                                    return DataSource.getGroupId(value.LoginName) ? false : true;
                                });

                                // Set the value
                                el.innerHTML = users.length.toString();
                            }

                            // Set the filter and order values
                            el.setAttribute("data-filter", el.innerHTML);
                            el.setAttribute("data-order", el.innerHTML);
                        }
                    },
                    {
                        name: "",
                        title: "Site Users",
                        onRenderCell: (el, col, item: IPermissionItem) => {
                            if (item.Type == "User") {
                                el.innerHTML = "1";
                            } else if (item.SiteMembers) {
                                // Get the users
                                let users = item.SiteMembers.filter((value, idx, self) => {
                                    return value.PrincipalType == SPTypes.PrincipalTypes.User && value.Email;
                                });

                                // Set the value
                                el.innerHTML = users.length.toString();
                            }

                            // Set the filter and order values
                            el.setAttribute("data-filter", el.innerHTML);
                            el.setAttribute("data-order", el.innerHTML);
                        }
                    },
                    {
                        name: "",
                        title: "Permission",
                        onRenderCell: (el, col, item: IPermissionItem) => {
                            // Parse the roles
                            for (let i = 0; i < item.Roles.length; i++) {
                                // Set the style
                                el.style.cursor = "pointer";

                                // Render a badge
                                let badge = Components.Badge({
                                    el,
                                    className: "me-2",
                                    isPill: true,
                                    content: item.Roles[i],
                                    type: Components.BadgeTypes.Primary
                                });

                                // Add a tooltip containing the text
                                Components.Tooltip({
                                    content: "<small>" + item.RoleInfo[i] + "</small>",
                                    target: badge.el
                                });
                            }
                        }
                    },
                    {
                        className: "text-end",
                        name: "",
                        title: "",
                        onRenderCell: (el, col, item: IPermissionItem) => {
                            // See if members exist
                            if (item.Type == "M365 Group" || item.Type == "Site Group") {
                                Components.Tooltip({
                                    el,
                                    content: "Click to view the members in this group.",
                                    btnProps: {
                                        text: "View Users",
                                        type: Components.ButtonTypes.OutlinePrimary,
                                        isSmall: true,
                                        onClick: () => {
                                            // View the users
                                            this.viewUsers(item);
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
        LoadingDialog.setHeader("Getting Site Roles");
        LoadingDialog.setBody("Searching the current permissions of the site...");
        LoadingDialog.show();

        // Clear the items
        this._groupIds = {};
        this._groupIdErrors = [];
        this._items = [];
        this._roleMapper = {};

        // Determine the webs to target
        let siteItems: Components.IDropdownItem[] = values["TargetWeb"] && values["TargetWeb"]["value"] ? [values["TargetWeb"]] as any : DataSource.SiteItems;

        // Parse all webs
        let counter = 0;
        Helper.Executor(siteItems, siteItem => {
            // Update the loading dialog
            LoadingDialog.setBody(`Getting the roles for web ${++counter} of ${siteItems.length}...`);

            // See if this is a sub-web and doesn't have unique role assignments
            // The sub-webs have the data property set for them
            if (siteItem.data != null && !siteItem.data.HasUniqueRoleAssignments) { return; }

            // Return a promise
            return new Promise(resolve => {
                // Get the permissions for the site
                Web(siteItem.text, { requestDigest: DataSource.SiteContext.FormDigestValue }).RoleAssignments().query({
                    Expand: [
                        "Member", "Member/Groups", "RoleDefinitionBindings"
                    ]
                }).execute(roles => {
                    // Update the loading dialog
                    LoadingDialog.setBody(`Analyzing the roles for web ${counter} of ${siteItems.length}...`);

                    // Analyze the roles
                    this.analyzeRoles(roles.results, siteItem.text, siteItem.data ? siteItem.data.Title : DataSource.Site.RootWeb.Title).then(resolve);
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

    // Views the users for an item
    private static viewUsers(item: IPermissionItem) {
        let rows: {
            Email: string;
            GroupType: "Member" | "Owner";
            Title: string;
            Type: "M365 Group" | "Site Group";
        }[] = [];

        // Set the modal header
        Modal.clear();
        Modal.setHeader("View Group Members/Owners");

        // Parse the site groups
        item.GroupMembers.forEach(groupUser => {
            // Add the row
            rows.push({
                Email: groupUser.mail,
                GroupType: "Member",
                Title: groupUser.displayName,
                Type: "M365 Group"
            });
        });
        item.GroupOwners.forEach(groupUser => {
            // Add the row
            rows.push({
                Email: groupUser.mail,
                GroupType: "Owner",
                Title: groupUser.displayName,
                Type: "M365 Group"
            });
        });

        // Parse the m365 groups
        item.SiteMembers.forEach(siteUser => {
            // Add the row
            rows.push({
                Email: siteUser.Email,
                GroupType: "Member",
                Title: siteUser.Title,
                Type: "Site Group"
            });
        });
        item.SiteOwners.forEach(siteUser => {
            // Add the row
            rows.push({
                Email: siteUser.Email,
                GroupType: "Owner",
                Title: siteUser.Title,
                Type: "Site Group"
            });
        });

        // Set the body
        new Dashboard({
            el: Modal.BodyElement,
            navigation: {
                title: "Users",
                showFilter: false
            },
            table: {
                rows,
                onRendering: dtProps => {
                    // Order by the 1st column then 2nd column
                    dtProps.order = [[0, "desc"], [1, "asc"]];

                    // Return the properties
                    return dtProps;
                },
                columns: [
                    {
                        name: "GroupType",
                        title: "User Type"
                    },
                    {
                        name: "Title",
                        title: "User Name"
                    },
                    {
                        name: "Email",
                        title: "User Email"
                    },
                    {
                        name: "Type",
                        title: "Group Containing User"
                    }
                ]
            }
        });

        // Set the footer
        Components.Button({
            el: Modal.FooterElement,
            text: "Close",
            onClick: () => {
                // Close the modal
                Modal.hide();
            }
        });

        // View the modal
        Modal.show();
    }
}