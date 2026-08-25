import { Dashboard, Documents, LoadingDialog, Modal } from "dattatable";
import { Components, ContextInfo, Helper, SPTypes, Types, Web } from "gd-sprest-bs";
import { DataSource } from "../ds";
import Strings from "../strings";
import { ExportCSV } from "./exportCSV";

interface ISearchItem {
    Email?: string;
    FileName?: string;
    FileUrl?: string;
    Group?: string;
    GroupId?: number;
    GroupInfo?: string;
    Id?: number;
    IsLimitedAccess: boolean;
    ItemId?: number;
    ListId?: string;
    ListName?: string;
    ListUrl?: string;
    LoginName: string;
    Name: string;
    ObjectType: "Site" | "Group" | "List" | "Item" | "File" | "User";
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
    "Name", "UserName", "Email", "ObjectType", "Group", "GroupInfo", "FileName", "FileUrl",
    "ItemId", "ListName", "ListUrl", "Role", "RoleInfo", "WebTitle", "WebUrl"
]

export class SearchEEEU {
    private static _dashboard: Dashboard = null;
    private static _elSubNav: HTMLElement = null;
    private static _items: ISearchItem[] = null;
    private static _loadOneDrive: boolean = null;
    private static _maxItemCount: number = 0;
    private static _oversharedGroups: string[] = null;
    private static _stopFl: boolean = false;

    // Analyzes a lists
    private static analyzeList(web: Types.SP.WebOData, list: Types.SP.ListOData, searchItems: boolean): PromiseLike<void> {
        // Return a promise
        return new Promise(resolve => {
            // Parse the role assignments
            Helper.Executor(list.RoleAssignments.results, roleAssignment => {
                let roleDef = (roleAssignment as any as Types.SP.RoleAssignmentOData).RoleDefinitionBindings.results[0];
                let user: Types.SP.User = roleAssignment.Member as any;

                // See if it's overshared
                if (this.isOvershared(user)) {
                    // Add a row for this entry
                    let roleItem: ISearchItem = {
                        Email: user.Email,
                        Group: "",
                        GroupId: 0,
                        GroupInfo: "",
                        Id: user.Id,
                        IsLimitedAccess: roleDef.Name == "Limited Access",
                        ListId: list.Id,
                        ListName: list.Title,
                        LoginName: user.LoginName,
                        ListUrl: list.RootFolder.ServerRelativeUrl,
                        Name: user.Title || user.LoginName,
                        ObjectType: "List",
                        Role: roleDef?.Name || "",
                        RoleInfo: roleDef?.Description || "",
                        WebUrl: web.Url,
                        WebTitle: web.Title
                    };
                    this._items.push(roleItem);
                    this._dashboard.Datatable.addRow(roleItem);
                }
            }).then(() => {
                if (searchItems) {
                    // Set the fields to query
                    let Select = ["Id", "HasUniqueRoleAssignments"];
                    if (list.BaseTemplate == SPTypes.ListTemplateType.DocumentLibrary ||
                        list.BaseTemplate == SPTypes.ListTemplateType.MySiteDocumentLibrary ||
                        list.BaseTemplate == SPTypes.ListTemplateType.PageLibrary) {
                        // Get the file information
                        Select.push("FileLeafRef");
                        Select.push("FileRef");
                    }

                    // Create a batch job
                    let completed = 0;
                    let ctrBatchJobs = 0;
                    let batchWeb = this._loadOneDrive ? Web.getOneDrive({ callbackQuery: () => { if (this._stopFl) { batch.stop(); } } })
                        : Web(web.Url, { requestDigest: DataSource.SiteContext.FormDigestValue, callbackQuery: () => { if (this._stopFl) { batch.stop(); } } });
                    let batch = batchWeb.Lists().getById(list.Id);

                    // Update the dialog
                    this._elSubNav.children[1].innerHTML = `Loading the items...`;

                    // Get the items for the list
                    let itemCounter = 0;
                    DataSource.loadItems({
                        isOnedrive: this._loadOneDrive,
                        listId: list.Id,
                        webUrl: web.Url,
                        query: { Select },
                        onItem: item => {
                            // Update the dialog
                            this._elSubNav.children[1].innerHTML = `Creating Batch Requests - Processed ${++itemCounter} items...`;

                            // See if this item doesn't have unique permissions
                            if (!item.HasUniqueRoleAssignments) { return this._stopFl; }

                            // Get the permissions
                            batch.Items(item.Id).RoleAssignments().query({
                                Expand: [
                                    "Member", "RoleDefinitionBindings"
                                ]
                            }).batch(roles => {
                                // Parse the role assignments
                                Helper.Executor(roles.results, roleAssignment => {
                                    let roleDef = roleAssignment.RoleDefinitionBindings.results[0];
                                    let user: Types.SP.User = roleAssignment.Member as any;

                                    // See if it's overshared
                                    if (this.isOvershared(user)) {
                                        // Add a row for this entry
                                        let roleItem: ISearchItem = {
                                            Email: user.Email,
                                            FileName: item["FileLeafRef"],
                                            FileUrl: item["FileRef"],
                                            Group: "",
                                            GroupId: 0,
                                            GroupInfo: "",
                                            Id: user.Id,
                                            IsLimitedAccess: roleDef.Name == "Limited Access",
                                            ItemId: item.Id,
                                            ListId: list.Id,
                                            ListName: list.Title,
                                            LoginName: user.LoginName,
                                            ListUrl: list.RootFolder.ServerRelativeUrl,
                                            Name: user.Title || user.LoginName,
                                            ObjectType: item["FileLeafRef"] ? "File" : "Item",
                                            Role: roleDef?.Name || "",
                                            RoleInfo: roleDef?.Description || "",
                                            WebUrl: web.Url,
                                            WebTitle: web.Title
                                        };
                                        this._items.push(roleItem);
                                        this._dashboard.Datatable.addRow(roleItem);
                                    }

                                    // Increment the counter and update the dialog
                                    this._elSubNav.children[1].innerHTML = `Batch Requests Processed ${++completed} of ${ctrBatchJobs % Strings.MaxBatchSize}...`;
                                });
                            }, ctrBatchJobs++ % Strings.MaxBatchSize == 0);

                            // Return the stop flag
                            return this._stopFl;
                        }
                    }).then(() => {
                        // Update the dialog
                        this._elSubNav.children[1].innerHTML = `Executing Batch Request for ${ctrBatchJobs} items...`;

                        // Execute the batch jobs
                        batch.execute(() => {
                            // Resolve the request
                            resolve();
                        });
                    });
                } else {
                    // Resolve the request
                    resolve();
                }
            });
        });
    }

    // Analyzes the site
    private static analyzeSite(web: Types.SP.WebOData, searchLists: boolean): PromiseLike<void> {
        // Return a promise
        return new Promise(resolve => {
            // Search the users
            this.analyzeUsers(web).then(() => {
                // Show a dialog
                this._elSubNav.children[1].innerHTML = `Analyzing lists...`;

                // Get the lists
                let site = this._loadOneDrive ? Web.getOneDrive() : Web(web.Url, { requestDigest: DataSource.SiteContext.FormDigestValue });
                site.Lists().query({
                    Filter: "Hidden eq false",
                    Expand: [
                        "RoleAssignments", "RoleAssignments/Groups", "RoleAssignments/Member",
                        "RoleAssignments/RoleDefinitionBindings", "RootFolder"
                    ],
                    Select: ["Id", "Title", "BaseTemplate", "HasUniqueRoleAssignments", "ItemCount", "RootFolder/ServerRelativeUrl"]
                }).execute(resp => {
                    let ctrList = 0;
                    let siteText = this._elSubNav.children[0].innerHTML;

                    // Parse the lists
                    let lists = this._loadOneDrive ? resp["value"] : resp.results;
                    Helper.Executor(lists, list => {
                        // See if we are stopping this process
                        if (this._stopFl) { return; }

                        // See if we are skipping large lists
                        if (this._maxItemCount > 0 && list.ItemCount > this._maxItemCount) { return; }

                        // Show a dialog
                        this._elSubNav.children[0].innerHTML = `${siteText} - [Analyzing List ${++ctrList} of ${lists.length}]: ${list.Title}`;

                        // Analyze the list
                        return this.analyzeList(web, list, searchLists);
                    }).then(() => {
                        // Resolve the request
                        resolve(null);
                    });
                });
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
                    // Update the dialog
                    this._elSubNav.children[1].innerHTML = `Analyzing User ${++counter} of ${users.length}`;

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
                name: "IncludeOversharedGroups",
                label: "Include Overshared Groups?",
                description: "Selecting this option will include the overshared groups in the search.",
                type: Components.FormControlTypes.Switch,
                value: true
            },
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
                let roleDef = role.RoleDefinitionBindings.results[0];

                // See if this role is the user
                if (role.Member.LoginName == userInfo.Name) {
                    // Add the user information
                    let roleItem: ISearchItem = {
                        WebUrl: web.Url,
                        WebTitle: web.Title,
                        Id: userInfo.Id,
                        IsLimitedAccess: roleDef.Name == "Limited Access",
                        LoginName: userInfo.Name,
                        Name: userInfo.Title || userInfo.Name,
                        ObjectType: "User",
                        Email: userInfo.EMail,
                        Group: "",
                        GroupId: 0,
                        GroupInfo: "",
                        Role: roleDef?.Name || "",
                        RoleInfo: roleDef?.Description || ""
                    };
                    this._items.push(roleItem);
                    this._dashboard.Datatable.addRow(roleItem);

                    // Check the next role
                    continue;
                }
            }

            // Get the groups the user is associated with
            let dstWeb = this._loadOneDrive ? Web.getOneDrive() : Web(DataSource.SiteContext.SiteFullUrl, { requestDigest: DataSource.SiteContext.FormDigestValue });
            dstWeb.SiteUsers(userInfo.Id).Groups().execute(groups => {
                // Parse the groups the member belongs to
                Helper.Executor(groups.results, group => {
                    // Parse the roles
                    for (let i = 0; i < web.RoleAssignments.results.length; i++) {
                        let role: Types.SP.RoleAssignmentOData = web.RoleAssignments.results[i] as any;

                        // See if the user belongs to this role
                        if (role.Member.LoginName == group.LoginName) {
                            let roleDef = role.RoleDefinitionBindings.results[0];

                            // Add the user information
                            let roleItem: ISearchItem = {
                                WebUrl: web.Url,
                                WebTitle: web.Title,
                                Id: userInfo.Id,
                                IsLimitedAccess: roleDef.Name == "Limited Access",
                                LoginName: userInfo.Name,
                                Name: userInfo.Title || userInfo.Name,
                                ObjectType: "Group",
                                Email: userInfo.EMail,
                                Group: group.Title,
                                GroupId: group.Id,
                                GroupInfo: group.Description || "",
                                Role: roleDef?.Name || "",
                                RoleInfo: roleDef?.Description || ""
                            };
                            this._items.push(roleItem);
                            this._dashboard.Datatable.addRow(roleItem);
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
            let web = this._loadOneDrive ? Web.getOneDrive() : Web(DataSource.SiteContext.SiteFullUrl, { requestDigest: DataSource.SiteContext.FormDigestValue });
            web.Lists("User Information List").Items().query({
                Select: ["Id", "Name", "EMail", "Title", "UserName"],
                GetAllItems: true,
                Top: 5000
            }).execute(items => {
                // Parse the items
                for (let i = 0; i < items.results.length; i++) {
                    let item = items.results[i];

                    // Add the user if this is flagged for overshared
                    if (this.isOvershared(item)) {
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
        });
    }

    // Returns true if the item is an overshared group
    private static isOvershared(item: Types.SP.ListItemOData | Types.SP.User): boolean {
        let isOvershared = false;

        // See if this is the default groups
        if (item.Title == "Everyone" || item.Title == "Everyone except external users" || item["Name"]?.indexOf("spo-grid-all-users") > 0) {
            // Set the flag
            isOvershared = true;
        }
        // Else, see if this is an overshared group
        else if (this._oversharedGroups.length > 0) {
            // Parse the overshared groups
            for (let j = 0; j < this._oversharedGroups.length; j++) {
                if (item.Title == this._oversharedGroups[j]) {
                    // Set the flag
                    isOvershared = true;
                    break;
                }
            }
        }

        // Return the flag
        return isOvershared;
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
    private static renderSummary(el: HTMLElement, auditOnly: boolean, items: ISearchItem[], onClose?: () => void) {

        // Create the nav items
        let navItems: Components.INavbarItem[] = onClose ? [{
            text: "New Search",
            className: "btn-outline-light",
            isButton: true,
            onClick: () => {
                // Set the stop flag
                this._stopFl = true;

                // Call the close event
                onClose();
            }
        }] : [];

        // Show the filter button for permissions
        navItems.push({
            text: "Show Limited Access",
            className: "btn-outline-light ms-2",
            isButton: true,
            onClick: () => {
                // Get the error button
                let elNav = el.querySelector("#navigation .navbar-nav");

                // Remove the last button
                let btn = elNav.querySelector("li:last-child a");

                // See if we are currently hiding limited access items
                if (btn.textContent == "Show Limited Access") {
                    // Remove the filter
                    this._dashboard.filter(2);

                    // Update the button text
                    btn.innerHTML = "Hide Limited Access";
                } else {
                    // Apply the filter
                    this._dashboard.filter(2, "false");

                    // Update the button text
                    btn.innerHTML = "Show Limited Access";
                }
            }
        });

        // Render the summary
        this._dashboard = new Dashboard({
            el,
            navigation: {
                title: "Search EEEU",
                showFilter: false,
                items: navItems,
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
                            "targets": 6,
                            "orderable": false,
                            "searchable": false
                        }
                    ];

                    // Hide the limited access column
                    dtProps.columnDefs.push({
                        "targets": 2,
                        "visible": false
                    });

                    // Order by the 1st column
                    dtProps.order = [[0, "asc"]];

                    // Return the properties
                    return dtProps;
                },
                columns: [
                    {
                        name: "Name",
                        title: "Information",
                        onRenderCell: (el, col, item: ISearchItem) => {
                            // Render the info
                            el.innerHTML = `
                                <b>Name: </b>${item.Name}
                                ${item.ListId ? `<br/><b>List: </b>${item.ListName}` : ""}
                                <br/>
                                <b>Web: </b>${item.WebUrl}
                            `;
                        }
                    },
                    {
                        name: "ObjectType",
                        title: "Object Type"
                    },
                    {
                        name: "IsLimitedAccess",
                        title: "Is<br/>Limited Access?"
                    },
                    {
                        name: "Group",
                        title: "Group Name"
                    },
                    {
                        name: "GroupInfo",
                        title: "Group Detail",
                        onRenderCell: (el) => {
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
                        name: "",
                        title: "Permissions",
                        onRenderCell: (el, col, item: ISearchItem) => {
                            // Set the style
                            el.style.cursor = "pointer";

                            // Render a badge
                            let badge = Components.Badge({
                                el,
                                className: "me-2",
                                isPill: true,
                                content: item.Role,
                                type: Components.BadgeTypes.Primary
                            });

                            // Add a tooltip containing the text
                            Components.Tooltip({
                                content: "<small>" + item.RoleInfo + "</small>",
                                target: badge.el
                            });
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

                            // Add the tooltip, based on the type
                            switch (row.ObjectType) {
                                case "File":
                                    // Add a button to the file
                                    tooltips.add({
                                        content: "Click to view the file.",
                                        btnProps: {
                                            className: "pe-2 py-1",
                                            text: "View File",
                                            type: Components.ButtonTypes.OutlinePrimary,
                                            onClick: () => {
                                                // View the file
                                                window.open(`${row.WebUrl}/${ContextInfo.layoutsUrl}/user.aspx`, "_blank");
                                            }
                                        }
                                    });
                                    break;
                                case "Group":
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
                                    break;
                                case "Item":
                                    // Add the view button
                                    tooltips.add({
                                        content: "Click to view the item unique permissions.",
                                        btnProps: {
                                            className: "pe-2 py-1",
                                            text: "View Group",
                                            type: Components.ButtonTypes.OutlinePrimary,
                                            onClick: () => {
                                                // View the group
                                                window.open(`${row.WebUrl}/${ContextInfo.layoutsUrl}/user.aspx?List=${row.ListId}&obj=${row.ListId},${row.FileUrl ? "1" : "2"},LISTITEM`, "_blank");
                                            }
                                        }
                                    });
                                    break;
                                case "List":
                                    // Add the view button
                                    tooltips.add({
                                        content: "Click to view the list permissions.",
                                        btnProps: {
                                            className: "pe-2 py-1",
                                            text: "View List",
                                            type: Components.ButtonTypes.OutlinePrimary,
                                            onClick: () => {
                                                // View the group
                                                window.open(`${row.WebUrl}/${ContextInfo.layoutsUrl}/user.aspx?List=${row.ListId}&obj=${row.ListId},LIST`, "_blank");
                                            }
                                        }
                                    });
                                    break;
                            }

                            // See if this is a list item
                            if (row.ListId) {
                            }

                            // Ensure this is a group
                            if (row.GroupId > 0) {
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
                                                if (confirm("Are you sure you want to revert the permissions to inherit from its parent?")) {
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

        // Set the sub-nav element
        this._elSubNav = el.querySelector("#sub-navigation");
        this._elSubNav.classList.remove("d-none");
        this._elSubNav.classList.add("my-2");
        this._elSubNav.innerHTML = `<div class="h6"></div><div></div>`;

        // Hide the limited access items by default
        this._dashboard.filter(2, "false");
    }

    // Runs the report
    static run(el: HTMLElement, auditOnly: boolean, values: { [key: string]: any }, onClose: () => void) {
        this._loadOneDrive = values["LoadOneDrive"] == "true";
        this._maxItemCount = parseInt(values["SkipLargeLists"]) || 0;
        this._stopFl = false;

        // Show a loading dialog
        LoadingDialog.setHeader("Searching Site");
        LoadingDialog.setBody("Searching the current permissions of the site...");
        LoadingDialog.show();

        // Clear the items
        this._items = [];

        // Set the overshared groups
        this._oversharedGroups = (values["IncludeOversharedGroups"] == true ? values["OversharedGroups"] : null) || [];

        // See if we are showing hidden lists
        let searchLists = values["SearchLists"];

        // Clear the element
        while (el.firstChild) { el.removeChild(el.firstChild); }

        // Render the summary
        this.renderSummary(el, auditOnly, this._items, onClose);

        // Hide the loading dialog
        LoadingDialog.hide();

        // Determine the webs to target
        let siteItems: Components.IDropdownItem[] = null;
        if (this._loadOneDrive) {
            siteItems = [{ text: DataSource.OneDriveWeb.Url, value: DataSource.OneDriveWeb.Id }] as any;
        } else {
            siteItems = values["TargetWeb"] && values["TargetWeb"]["value"] ? [values["TargetWeb"]] as any : DataSource.SiteItems;
        }

        // Parse all webs
        let counter = 0;
        Helper.Executor(siteItems, siteItem => {
            // See if we are stopping this process
            if (this._stopFl) { return; }

            // Update the status
            this._elSubNav.children[0].innerHTML = `Searching Site ${++counter} of ${siteItems.length}`;
            this._elSubNav.children[1].innerHTML = "Getting the info for the web...";

            // Return a promise
            return new Promise(resolve => {
                // Get the permissions
                let web = this._loadOneDrive ? Web.getOneDrive() : Web(siteItem.text, { requestDigest: DataSource.SiteContext.FormDigestValue });
                web.query({
                    Expand: [
                        "RoleAssignments", "RoleAssignments/Groups", "RoleAssignments/Member",
                        "RoleAssignments/RoleDefinitionBindings", "SiteGroups"
                    ]
                }).execute(web => {
                    // Update the dialog
                    this._elSubNav.children[1].innerHTML = `Analyzing web ${counter} of ${siteItems.length}...`;

                    // Analyze the site
                    this.analyzeSite(web, searchLists).then(resolve);
                });
            });
        }).then(() => {
            // Hide the sub-nav
            this._elSubNav.classList.add("d-none");
        });
    }

    // Searches a list for EEEU
    static searchList(webUrl: string, listName: string, auditOnly: boolean, oversharedGroups: string[] = []) {
        this._loadOneDrive = false;
        this._stopFl = false;

        // Clear the items
        this._items = [];

        // Set the overshared groups
        this._oversharedGroups = oversharedGroups;

        // Clear the modal
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

        // Render the summary
        this.renderSummary(Modal.BodyElement, auditOnly, this._items);

        // Show the modal
        Modal.show();

        // Update the status
        this._elSubNav.children[0].innerHTML = `Searching List: ${listName}`;
        this._elSubNav.children[1].innerHTML = "Getting the info for the web...";

        // Get the permissions
        let web = Web(webUrl, { requestDigest: DataSource.SiteContext.FormDigestValue });
        web.query({
            Expand: [
                "RoleAssignments", "RoleAssignments/Groups", "RoleAssignments/Member",
                "RoleAssignments/RoleDefinitionBindings", "SiteGroups"
            ]
        }).execute(webInfo => {
            // Update the dialog
            this._elSubNav.children[1].innerHTML = `Analyzing the list...`;

            // Get the list information
            web.Lists(listName).query({
                Expand: [
                    "RoleAssignments", "RoleAssignments/Groups", "RoleAssignments/Member",
                    "RoleAssignments/RoleDefinitionBindings", "RootFolder"
                ],
                Select: ["Id", "Title", "BaseTemplate", "HasUniqueRoleAssignments", "RootFolder/ServerRelativeUrl"]
            }).execute(list => {
                // Analyze the list
                this.analyzeList(webInfo, list, true).then(() => {
                    // Hide the sub-nav
                    this._elSubNav.classList.add("d-none");
                });
            });
        });
    }

    // Stops the report
    static stop() { this._stopFl = true; }
}