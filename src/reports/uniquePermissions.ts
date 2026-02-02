import { Dashboard, Documents, LoadingDialog } from "dattatable";
import { Components, Helper, SPTypes, Types, Web } from "gd-sprest-bs";
import { cardList } from "gd-sprest-bs/build/icons/svgs/cardList";
import { DataSource } from "../ds";
import { ExportCSV } from "./exportCSV";

interface IPermission {
    FileName?: string;
    FileUrl?: string;
    ItemId: number;
    ListId: string;
    ListName: string;
    ListType: number;
    ListUrl: string;
    ListViewUrl: string;
    RoleAssignmentId?: number;
    SiteGroupId: number;
    SiteGroupName: string;
    SiteGroupPermission: string;
    SiteGroupUrl?: string;
    SiteGroupUsers?: string;
    WebUrl: string;
}

const CSVFields = [
    "WebUrl", "ListType", "ListName", "ListUrl", "ItemId", "FileName", "FileUrl", "Permissions", "RoleAssignmentId",
    "SiteGroupId", "SiteGroupName", "SiteGroupPermission", "SiteGroupUrl", "SiteGroupUsers"
]

export class UniquePermissions {
    private static _dashboard: Dashboard = null;
    private static _elSubNav: HTMLElement = null;
    private static _items: IPermission[] = [];

    // Analyzes a list
    private static analyzeList(webUrl: string, list: Types.SP.ListOData): PromiseLike<void> {
        // Return a promise
        return new Promise(resolve => {
            let Select = ["Id", "HasUniqueRoleAssignments"];

            // See if this is a document library
            if (list.BaseTemplate == SPTypes.ListTemplateType.DocumentLibrary || list.BaseTemplate == SPTypes.ListTemplateType.PageLibrary) {
                // Get the file information
                Select.push("FileLeafRef");
                Select.push("FileRef");
            }

            // Update the dialog
            this._elSubNav.children[1].innerHTML = "Getting the list items...";

            // Get the items where it has broken inheritance
            Web(webUrl, { requestDigest: DataSource.SiteContext.FormDigestValue }).Lists().getById(list.Id).Items().query({
                GetAllItems: true,
                Select,
                Top: 5000
            }).execute(items => {
                let ctrBatchJobs = 0;

                // Create a batch job
                let batch = Web(webUrl, { requestDigest: DataSource.SiteContext.FormDigestValue }).Lists().getById(list.Id);

                // Update the dialog
                this._elSubNav.children[1].innerHTML = "Creating the batch job...";

                // Parse the items
                Helper.Executor(items.results, item => {
                    // See if this item doesn't have unique permissions
                    if (!item.HasUniqueRoleAssignments) { return; }

                    // Get the permissions
                    batch.Items(item.Id).RoleAssignments().query({
                        Expand: [
                            "Member/Users", "RoleDefinitionBindings"
                        ]
                    }).batch(roles => {
                        // Parse the role assignments
                        Helper.Executor(roles.results, roleAssignment => {
                            // Parse the role definitions and create a list of permissions
                            let roleDefinitions = [];
                            for (let i = 0; i < roleAssignment.RoleDefinitionBindings.results.length; i++) {
                                // Add the permission name
                                roleDefinitions.push(roleAssignment.RoleDefinitionBindings.results[i].Name);
                            }

                            // See if this is a group
                            if (roleAssignment.Member["Users"] != null) {
                                let group: Types.SP.GroupOData = roleAssignment.Member as any;

                                // Parse the users and create a list of members
                                let members = [];
                                for (let i = 0; i < group.Users.results.length; i++) {
                                    let user = group.Users.results[i];

                                    // Add the user information
                                    members.push(user.Email || user.UserPrincipalName || user.Title);
                                }

                                // Add a row for this entry
                                let listItem = {
                                    FileName: item["FileLeafRef"],
                                    FileUrl: item["FileRef"],
                                    ItemId: item.Id,
                                    ListId: list.Id,
                                    ListName: list.Title,
                                    ListType: list.BaseTemplate,
                                    ListUrl: list.RootFolder.ServerRelativeUrl,
                                    ListViewUrl: list.DefaultDisplayFormUrl,
                                    SiteGroupId: group.Id,
                                    SiteGroupName: group.LoginName,
                                    SiteGroupPermission: roleDefinitions.join(', '),
                                    SiteGroupUrl: DataSource.SiteContext.SiteFullUrl + "/_layouts/15/people.aspx?MembershipGroupId=" + group.Id,
                                    SiteGroupUsers: members.join(', '),
                                    WebUrl: webUrl
                                };
                                this._items.push(listItem);
                                this._dashboard.Datatable.addRow(listItem);
                            } else {
                                let user: Types.SP.User = roleAssignment.Member as any;

                                // Add a row for this entry
                                let listItem = {
                                    FileName: item["FileLeafRef"],
                                    FileUrl: item["FileRef"],
                                    ItemId: item.Id,
                                    ListId: list.Id,
                                    ListName: list.Title,
                                    ListType: list.BaseTemplate,
                                    ListUrl: list.RootFolder.ServerRelativeUrl,
                                    ListViewUrl: list.DefaultDisplayFormUrl,
                                    RoleAssignmentId: roleAssignment.PrincipalId,
                                    SiteGroupId: user.Id,
                                    SiteGroupName: user.Email || user.UserPrincipalName || user.Title,
                                    SiteGroupPermission: roleDefinitions.join(', '),
                                    WebUrl: webUrl
                                };
                                this._items.push(listItem);
                                this._dashboard.Datatable.addRow(listItem);
                            }
                        });
                    }, ctrBatchJobs++ % 25 == 0);
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

    // Gets the form fields to display
    static getFormFields(): Components.IFormControlProps[] { return []; }

    // Renders the search summary
    private static renderSummary(el: HTMLElement, auditOnly: boolean, onClose: () => void) {
        // Render the summary
        this._dashboard = new Dashboard({
            el,
            navigation: {
                title: "Unique Permissions",
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
                        new ExportCSV("uniquePermissions.csv", CSVFields, this._items);
                    }
                }]
            },
            table: {
                rows: this._items,
                onRendering: dtProps => {
                    dtProps.columnDefs = [
                        {
                            "targets": 6,
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
                        name: "ListName",
                        title: "List Name"
                    },
                    {
                        name: "FileName",
                        title: "File Name"
                    },
                    {
                        name: "SiteGroupName",
                        title: "Site Group"
                    },
                    {
                        name: "SiteGroupPermission",
                        title: "Permission"
                    },
                    {
                        name: "SiteGroupUsers",
                        title: "User(s)"
                    },
                    {
                        className: "text-end",
                        name: "",
                        title: "",
                        onRenderCell: (el, col, item: IPermission) => {
                            let btnDelete: Components.IButton = null;

                            // Render the buttons
                            Components.TooltipGroup({
                                el,
                                tooltips: [
                                    {
                                        content: "Click to view the item properties.",
                                        btnProps: {
                                            className: "pe-2 py-1",
                                            iconClassName: "mx-1",
                                            iconSize: 24,
                                            iconType: cardList,
                                            text: "View",
                                            type: Components.ButtonTypes.OutlinePrimary,
                                            onClick: () => {
                                                // Show the security group
                                                window.open(item.ListViewUrl + "?ID=" + item.ItemId, "_blank");
                                            }
                                        }
                                    },
                                    {
                                        content: "Click to view the file.",
                                        className: item.FileName ? "" : "d-none",
                                        btnProps: {
                                            className: "pe-2 py-1",
                                            //iconType: GetIcon(24, 24, "CustomList", "mx-1"),
                                            text: "File",
                                            type: Components.ButtonTypes.OutlinePrimary,
                                            onClick: () => {
                                                // Open the document in view mode
                                                Documents.open(item.FileUrl, false, item.WebUrl);
                                            }
                                        }
                                    },
                                    {
                                        content: "Click to restore permissions to inherit.",
                                        btnProps: {
                                            assignTo: btn => { btnDelete = btn; },
                                            className: "pe-2 py-1",
                                            //iconType: GetIcon(24, 24, "PeopleTeamDelete", "mx-1"),
                                            text: "Restore",
                                            type: Components.ButtonTypes.OutlinePrimary,
                                            onClick: () => {
                                                // Confirm the deletion of the group
                                                if (confirm("Are you sure you restore the permissions to inherit?")) {
                                                    // Revert the permissions
                                                    this.revertPermissions(item);
                                                }
                                            }
                                        }
                                    }
                                ]
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
    private static revertPermissions(item: IPermission) {
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

    // Runs the report
    static run(el: HTMLElement, auditOnly: boolean, values: { [key: string]: string }, onClose: () => void) {
        // Show a loading dialog
        LoadingDialog.setHeader("Searching Lists");
        LoadingDialog.setBody("Searching the site...");
        LoadingDialog.show();

        // Clear the items
        this._items = [];

        // Clear the element
        while (el.firstChild) { el.removeChild(el.firstChild); }

        // Render the summary
        this.renderSummary(el, auditOnly, onClose);

        // Hide the loading dialog
        LoadingDialog.hide();

        // Determine the webs to target
        let siteItems: Components.IDropdownItem[] = values["TargetWeb"] && values["TargetWeb"]["value"] ? [values["TargetWeb"]] as any : DataSource.SiteItems;

        // Parse all the webs
        let counter = 0;
        Helper.Executor(siteItems, siteItem => {
            // Update the status
            let siteText = `Searching Site ${++counter} of ${siteItems.length}`;
            this._elSubNav.children[0].innerHTML = siteText;
            this._elSubNav.children[1].innerHTML = "Loading the lists for this site...";

            // Return a promise
            return new Promise(resolve => {
                // Get the lists
                Web(siteItem.text, { requestDigest: DataSource.SiteContext.FormDigestValue }).Lists().query({
                    Filter: "Hidden eq false",
                    Expand: ["DefaultDisplayFormUrl", "DefaultViewFormUrl", "RootFolder"],
                    Select: ["BaseTemplate", "Id", "Title", "HasUniqueRoleAssignments", "RootFolder/ServerRelativeUrl"]
                }).execute(lists => {
                    let ctrList = 0;

                    // Update the dialog
                    this._elSubNav.children[1].innerHTML = "Loading the lists for this site...";

                    // Parse the lists
                    Helper.Executor(lists.results, list => {
                        // Update the dialog
                        this._elSubNav.children[0].innerHTML = `${siteText} [Analyzing Library ${++ctrList} of ${lists.results.length}]: ${list.Title}`;
                        this._elSubNav.children[1].innerHTML = "Analyzing the list...";

                        // Analyze the list
                        return this.analyzeList(siteItem.text, list);
                    }).then(resolve);
                });
            });
        }).then(() => {
            // Hide the sub-nav
            this._elSubNav.classList.add("d-none");
        });
    }
}