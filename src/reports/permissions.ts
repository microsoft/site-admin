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
    private static _items: IPermissionItem[] = null;

    // Analyze the group ids
    private static analyzeGroupIds(groupIds: string[]): PromiseLike<void> {
        // Return a promise
        return new Promise(resolve => {
            let validGroupIds = [];

            // Try to get the groups
            let counter = 0;
            Helper.Executor(groupIds, groupId => {
                // Update the loading dialog
                LoadingDialog.setBody(`Trying to get M365 Group information ${++counter} of ${groupIds.length}`);

                // Return a promise
                return new Promise(resolve => {
                    // Get the group information
                    DirectorySession().group(groupId).execute(group => {
                        // Append the group
                        validGroupIds.push(group.id);
                    }, () => {
                        // Try the next group
                        resolve(null);
                    });
                });
            }).then(() => {
                // Update the loading dialog
                LoadingDialog.setBody(`Creating the batch request to get M365 Group information`);

                // Create the request
                let ds = DirectorySession();

                // Parse the group ids
                groupIds.forEach(groupId => {
                    // Get the M365 group information
                    ds.group(groupId).query({
                        Expand: ["members", "owners"],
                        Select: [
                            "calendarUrl", "displayName", "id", "isPublic", "mail",
                            "members/principalName", "members/id", "members/displayName", "members/mail",
                            "owners/principalName", "owners/id", "owners/displayName", "owners/mail"
                        ]
                    }).batch(group => {
                        // Ensure the group was found
                        if (group.id) {
                            // Add it to the mapper
                            this._groupIds[group.id] = group;
                        }
                    });
                });

                // Update the loading dialog
                LoadingDialog.setBody(`Executing the batch request to get M365 Group information`);

                // Execute the batch job
                ds.execute(() => {
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
        });
    }

    // Analyze the role
    private static analyzeRole(role: Types.SP.RoleAssignmentOData): string[] {
        let item: IPermissionItem = null;

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
                WebTitle: DataSource.Site.RootWeb.Title,
                WebUrl: DataSource.Site.Url,
                Roles: [],
                RoleInfo: [],
                Type: "User"
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
                SiteMembers: role.Member["Users"].results,
                SiteOwners: [role.Member["Owner"]],
                Type: "Site Group",
                WebTitle: DataSource.Site.RootWeb.Title,
                WebUrl: DataSource.Site.Url,
            };
        } else {
            // See if this is a M365 group
            let groupId = this.getGroupId(role.Member.LoginName);

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
                SiteMembers: role.Member["Users"]?.results || [],
                SiteOwners: role.Member["Owner"] ? [role.Member["Owner"]] : [],
                Type: groupId ? "M365 Group" : "AD Group",
                WebTitle: DataSource.Site.RootWeb.Title,
                WebUrl: DataSource.Site.Url
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

        // Return the group ids
        return item.GroupIds;
    }

    // Analyzes the roles
    private static analyzeRoles(roles: Types.SP.RoleAssignmentOData[]): PromiseLike<void> {
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

                // Analyze the role
                groupIds = groupIds.concat(this.analyzeRole(role));
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
                // Add the group id
                let groupId = this.getGroupId(user.LoginName);
                groupId ? item.GroupIds.push(groupId) : null;
            }
        }
    }

    // Gets the form fields to display
    static getFormFields(): Components.IFormControlProps[] { return []; }

    // Returns the group id from the login name
    static getGroupId(loginName: string): string {
        // Get the group id from the login name
        let userInfo = loginName.split('|');
        let groupInfo = userInfo[userInfo.length - 1];
        let groupId = groupInfo.split('_')[0];

        // Ensure it's a guid and return null if it's not
        return /^[{]?[0-9a-fA-F]{8}(-[0-9a-fA-F]{4}){3}-[0-9a-fA-F]{12}[}]?$/.test(groupId) ? groupId : null;
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
    private static renderSummary(el: HTMLElement, items: IPermissionItem[], onClose: () => void) {
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
                        // Need to generate the CSV for this.....
                        new ExportCSV("searchDocs.csv", CSVFields, items);
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
                        title: "Site Users",
                        onRenderCell: (el, col, item: IPermissionItem) => {
                            if (item.Type == "User") {
                                el.innerHTML = "1";
                            } else {
                                // Get the members and owners and filter for duplicates
                                let users = (item.SiteMembersAsString.concat(item.SiteOwnersAsString)).filter((value, idx, self) => {
                                    return self.indexOf(value) === idx;
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
                            if (item.Type == "M365 Group") {
                                // Set the value
                                el.innerHTML = (item.GroupMembersAsString.length + item.GroupOwnersAsString.length).toString();
                            } else {
                                el.innerHTML = "1";
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
    static run(el: HTMLElement, values: { [key: string]: any }, onClose: () => void) {
        // Show a loading dialog
        LoadingDialog.setHeader("Searching Site");
        LoadingDialog.setBody("Searching the current permissions of the site...");
        LoadingDialog.show();

        // Clear the items
        this._groupIds = {};
        this._items = [];

        // Get the permissions
        Web(DataSource.SiteContext.SiteFullUrl, { requestDigest: DataSource.SiteContext.FormDigestValue }).RoleAssignments().query({
            Expand: [
                "Member", "Member/Groups", "Member/Owner", "Member/Users", "RoleDefinitionBindings"
            ]
        }).execute(roles => {
            // Analyze the roles
            this.analyzeRoles(roles.results).then(() => {
                // Clear the element
                while (el.firstChild) { el.removeChild(el.firstChild); }

                // Render the summary
                this.renderSummary(el, this._items, onClose);

                // Hide the loading dialog
                LoadingDialog.hide();
            });
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