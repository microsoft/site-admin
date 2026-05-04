import { CanvasForm, Dashboard, LoadingDialog } from "dattatable";
import { Components, ContextInfo, SPTypes, Types, Web } from "gd-sprest-bs";
import { DataSource } from "../ds";
import { M365Groups } from "../m365Groups";
import { IDLPItem } from "./dlp";
import { ISensitivityLabelItem } from "./sensitivityLabels";

interface IPermission {
    GroupName?: string;
    Member: string;
    Permission: string;
    Type: string;
}

/**
 * View Permissions Report
 */
export class ViewPermissions {
    private static _oversharedGroups: string[] = [];
    static get OversharedGroups(): string[] { return this._oversharedGroups; }
    static set OversharedGroups(value: string[]) { this._oversharedGroups = value; }

    // Returns the groups that are flagging the file as overshared
    static getOversharedGroups(roles: Types.SP.RoleAssignmentOData[]): string[] {
        let groups = [];

        // Check if any role assignment matches the overshared groups
        for (let i = 0; i < roles.length; i++) {
            let role = roles[i];

            // See if this is the eeeu, everyone or the custom group
            if (role.Member.Title == "Everyone except external users" || role.Member.Title == "Everyone" || this.OversharedGroups.indexOf(role.Member.Title) >= 0) {
                // Append the group
                groups.push(role.Member.Title);
            }
            // Else, go through the users
            else {
                let users: Types.SP.User[] = role.Member["Users"] ? role.Member["Users"].results : [];
                users.forEach(user => {
                    // See if this is the eeeu or everyone
                    if (user.Title == "Everyone except external users" || user.Title == "Everyone" || this.OversharedGroups.indexOf(user.Title) >= 0) {
                        // Append the group
                        groups.push("<b>" + role.Member.Title + "</b> - " + user.Title);
                    }
                });
            }
        }

        // Return the groups
        return groups;
    }

    // Determines if a file is overshared
    static isOvershared(roles: Types.SP.RoleAssignmentOData[]): boolean {
        let isOvershared = false;

        // Check if any role assignment matches the overshared groups
        for (let i = 0; i < roles.length; i++) {
            let role = roles[i];

            // See if this is the eeeu or everyone
            if (role.Member.Title == "Everyone except external users" || role.Member.Title == "Everyone" || this.OversharedGroups.indexOf(role.Member.Title) >= 0) {
                // Set the flag
                isOvershared = true;
                break;
            }
            // Else, see if it's one of the custom groups
            else {
                // Parse the users
                let users: Types.SP.User[] = role.Member["Users"] ? role.Member["Users"].results : [];
                for (let i = 0; i < users.length; i++) {
                    let user = users[i];

                    // See if this is the eeeu or everyone
                    if (user.Title == "Everyone except external users" || user.Title == "Everyone" || this.OversharedGroups.indexOf(user.Title) >= 0) {
                        // Set the flag
                        isOvershared = true;
                        break;
                    }
                }

                // Break from the loop if the flag is set
                if (isOvershared) { break; }
            }
        }

        // Return the flag
        return isOvershared;
    }

    // Removes the groups that are flagging the file as overshared
    static removeOversharedGroups(item: IDLPItem | ISensitivityLabelItem, onComplete: () => void) {
        // Show a canvas form
        CanvasForm.clear();
        CanvasForm.setHeader("Secure File");
        CanvasForm.setType(Components.OffcanvasTypes.End);
        CanvasForm.setSize(Components.OffcanvasSize.Small2);

        // Set the content
        CanvasForm.setBody(`
            <p>This action will create unique permissions for the file and remove explicit permissions to groups that are flagging this file as overshared.</p>
            <p>If access to the file is currently granted through a SharePoint group, then that group will be removed from the file's permissions. For example, if the "Visitors Group" contains the "All DoD Guests" group; then the "Visitors Group" will be removed from the file's permissions.</p>
            <p>It is recommended to review the file's permissions after securing it to ensure it is properly permissioned to the people who require access.</p>
            <p>The following groups will be removed from the file's permissions:</p>
            <ul></ul>
            <div class="d-flex justify-content-end"></div>
        `);

        // Render the groups that are flagging this file as overshared
        let ul = CanvasForm.BodyElement.querySelector("ul");
        this.getOversharedGroups(item.Permissions).forEach(group => {
            // Add the group
            let li = document.createElement("li");
            li.innerHTML = group;
            ul.appendChild(li);
        });

        // Render the footer
        Components.ButtonGroup({
            el: CanvasForm.BodyElement.querySelector("div"),
            className: "mt-3",
            buttons: [
                {
                    text: "Confirm",
                    type: Components.ButtonTypes.OutlinePrimary,
                    onClick: () => {
                        // Hide the canvas form
                        CanvasForm.hide();

                        // Show a loading dialog
                        LoadingDialog.setHeader("Clearing Permissions");
                        LoadingDialog.setBody("Breaking inheritance for this item...");
                        LoadingDialog.show();

                        // Set the list containing the item
                        let list = Web(item.WebUrl, {
                            disableCache: true,
                            requestDigest: DataSource.SiteContext.FormDigestValue
                        }).Lists().getById(item.ListId);

                        // See if we need to break inheritance
                        if (!item.HasUniquePermissions) {
                            // Break role inheritance and copy the permissions
                            list.Items(item.ItemId).breakRoleInheritance(true, false).execute(() => {
                                // Update the dialog
                                LoadingDialog.setBody("Getting the permissions for this item...");
                            });
                        }

                        // Get the role assignments for this item
                        list.Items(item.ItemId).RoleAssignments().query({ Expand: ["Member/Users"] }).execute(permissions => {
                            let roles = list.Items(item.ItemId).RoleAssignments();

                            // Update the dialog
                            LoadingDialog.setBody("Removing the groups for this...");

                            // Check if any role assignment matches the overshared groups
                            for (let i = 0; i < permissions.results.length; i++) {
                                let permission = permissions.results[i];

                                // See if this is the eeeu, everyone or the custom group
                                if (permission.Member.Title == "Everyone except external users" || permission.Member.Title == "Everyone" || this.OversharedGroups.indexOf(permission.Member.Title) >= 0) {
                                    // Remove the role assignment
                                    roles.removeRoleAssignment(permission.PrincipalId).execute(true);
                                }
                                // Else, go through the users
                                else {
                                    let users: Types.SP.User[] = permission.Member["Users"] ? permission.Member["Users"].results : [];
                                    users.forEach(user => {
                                        // See if this is the eeeu or everyone
                                        if (user.Title == "Everyone except external users" || user.Title == "Everyone" || this.OversharedGroups.indexOf(user.Title) >= 0) {
                                            // Remove the role assignment
                                            roles.removeRoleAssignment(permission.PrincipalId).execute(true);
                                        }
                                    });
                                }
                            }

                            // Wait for the requests to complete
                            roles.done(() => {
                                // Call the complete event
                                onComplete();

                                // Hide the loading dialog
                                LoadingDialog.hide();
                            });
                        }, true);
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

    // Views the permissions for the item
    static show(item: IDLPItem | ISensitivityLabelItem) {
        // Clear the canvas form
        CanvasForm.clear();
        CanvasForm.setHeader("View Permissions");
        CanvasForm.setSize(Components.OffcanvasSize.Large2);
        CanvasForm.setType(Components.OffcanvasTypes.End);

        // Set the row data
        let rows: IPermission[] = [];

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
                    let groupId = M365Groups.getGroupId(role.Member.LoginName);
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
                onRendering(dtProps) {
                    // Set the column defs
                    dtProps.columnDefs = [
                        {
                            "targets": [4],
                            "orderable": false,
                            "searchable": false
                        }
                    ];

                    // Return the properties
                    return dtProps;
                },
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
                    },
                    {
                        name: "",
                        title: "Actions",
                        onRenderCell: (el, col, row: IPermission) => {
                            let tooltips: Components.ITooltipProps[] = [];

                            // See if this is a site group and not the limited access group
                            if (row.Type === "Site Group") {
                                // Set the url
                                let url = `${item.WebUrl}/${ContextInfo.layoutsUrl}/people.aspx?MembershipGroupId=${item.ItemId}`;
                                if (row.GroupName.indexOf("Limited Access System Group") === 0) {
                                    // Update the url to the limited access group
                                    url = `${item.WebUrl}/${ContextInfo.layoutsUrl}/user.aspx?List=${item.ListId}&obj=${item.ListId},${item.ItemId},LISTITEM&showLimitedAccessUsers=true`;
                                }

                                tooltips.push({
                                    content: "Click to view the site group.",
                                    btnProps: {
                                        text: "View Group",
                                        type: Components.ButtonTypes.OutlinePrimary,
                                        onClick: () => {
                                            window.open(url, "_blank");
                                        }
                                    }
                                });
                            }

                            // See if this is an ad group or user
                            if (row.Type === "AD Group" || row.Type === "User") {
                                tooltips.push({
                                    content: "Click to view the user.",
                                    btnProps: {
                                        text: row.Type === "AD Group" ? "View Group" : "View User",
                                        type: Components.ButtonTypes.OutlinePrimary,
                                        onClick: () => {
                                            window.open(`${item.WebUrl}/${ContextInfo.layoutsUrl}/userdisp.aspx?ID=${item.ItemId}`, "_blank");
                                        }
                                    }
                                });
                            }

                            // Render the actions
                            Components.TooltipGroup({
                                el,
                                isVertical: true,
                                tooltips
                            });
                        }
                    }
                ]
            }
        });

        // Show the form
        CanvasForm.show();
    }
}