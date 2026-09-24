import { Dashboard, Modal } from "dattatable";
import { Components, SPTypes, Web } from "gd-sprest-bs";
import { IAppProps } from "../app";
import { DataSource, ISiteUserInfo } from "../ds";

/**
 * Access
 */
export class AccessTab {
    private _appProps: IAppProps = null;
    private _admins: ISiteUserInfo[] = null;
    private _el: HTMLElement = null;
    private _owners: ISiteUserInfo[] = null;
    private _webUrl: string = null;

    // Constructor
    constructor(el: HTMLElement, appProps: IAppProps) {
        this._appProps = appProps;
        this._el = el;
        this._webUrl = DataSource.Web.Url;

        // Render the tab
        this.render();
    }

    // Shows the add user form
    private addUser(onAddUser: () => void) {
        // Clear the modal
        Modal.clear();

        // Set the header
        Modal.setHeader("Add User");

        // Determine the dropdown options based on the user's current permission
        let items: Components.IDropdownItem[] = [{ text: "Owner", value: "Owner" }];
        if (DataSource.IsAdmin) { items.push({ text: "Admin", value: "Admin" }); }

        // Set the form
        let form = Components.Form({
            el: Modal.BodyElement,
            controls: [
                {
                    name: "permission",
                    type: Components.FormControlTypes.Dropdown,
                    label: "Permission",
                    required: true,
                    items
                } as Components.IFormControlPropsDropdown,
                {
                    name: "user",
                    type: Components.FormControlTypes.PeoplePicker,
                    label: "User",
                    description: "Select the user to add to the site.",
                    required: true,
                    onValidate: (ctrl, results) => {
                        // Default the validation
                        results.isValid = true;

                        // Ensure a user is selected
                        let user = results.value[0];
                        if (user) {
                            // Ensure an email exists
                            if ((user.Email || "").length === 0) {
                                // Set the error message
                                results.isValid = false;
                                results.invalidMessage = "The selected user account does not have an email address associated with it.";

                            } else {
                                // See if we are adding an admin
                                if (form.getControl("permission").getValue().value === "Admin") {
                                    // Parse the admins
                                    for (let i = 0; i < this._admins.length; i++) {
                                        if (this._admins[i].email === user.Email) {
                                            results.isValid = false;
                                            results.invalidMessage = "The selected user is already an admin.";
                                            break;
                                        }
                                    }
                                } else {
                                    // Parse the owners
                                    for (let i = 0; i < this._owners.length; i++) {
                                        if (this._owners[i].email === user.Email) {
                                            results.isValid = false;
                                            results.invalidMessage = "The selected user is already an owner.";
                                            break;
                                        }
                                    }
                                }
                            }
                        } else {
                            // Set the error message
                            results.isValid = false;
                            results.invalidMessage = "The user selection is required.";
                        }

                        // Return the results
                        return results;
                    }
                } as Components.IFormControlPropsPeoplePicker
            ]
        });

        // Set the footer
        Components.TooltipGroup({
            el: Modal.FooterElement,
            tooltips: [
                {
                    content: "Click to add the user.",
                    btnProps: {
                        text: "Add",
                        onClick: () => {
                            // Ensure the form is valid
                            if (!form.isValid()) { return; }

                            // Get the user email
                            let values = form.getValues();
                            let permission = values["permission"].value;
                            let userEmail = values["user"][0].Email;

                            // Set the web
                            let web = Web(this._webUrl, { requestDigest: DataSource.SiteContext.FormDigestValue });

                            // Ensure the user exists in the site
                            web.ensureUser(userEmail).execute((user) => {
                                // See if this is an admin or owner
                                if (permission === "Admin") {
                                    // Set the flag
                                    user.update({ IsSiteAdmin: true }).execute(() => {
                                        // Add the user to the admin
                                        this._admins.push({
                                            email: user.Email,
                                            id: user.Id,
                                            name: user.LoginName,
                                            permission: "Admin",
                                            title: user.Title,
                                            type: user.PrincipalType
                                        });

                                        // Call the event
                                        onAddUser();
                                    });
                                } else {
                                    // Add the user to the default owner's group
                                    web.AssociatedOwnerGroup().Users().addUserById(user.Id).execute(() => {
                                        // Add the user to the admin
                                        this._admins.push({
                                            email: user.Email,
                                            id: user.Id,
                                            name: user.LoginName,
                                            permission: "Owner",
                                            title: user.Title,
                                            type: user.PrincipalType
                                        });

                                        // Call the event
                                        onAddUser();
                                    });
                                }
                            }, () => {
                                // Error adding the user
                                console.error("Error adding the user to the site. Refresh the page and try again.");

                                // Call the event
                                onAddUser();
                            });
                        }
                    }
                },
                {
                    content: "Click to close the form.",
                    btnProps: {
                        text: "Close",
                        onClick: () => {
                            Modal.hide();
                        }
                    }
                }
            ]
        });

        // Show the modal
        Modal.show();
    }

    // Returns the permissions for the admins/owners
    private loadUsers(): PromiseLike<void> {
        // Return a promise
        return new Promise(resolve => {
            let ctr = 0;

            // See if we need to load the admins
            if (this._admins === null) {
                // Load the site admins
                DataSource.loadSiteAdministrators().then(admins => {
                    // Set the admins
                    this._admins = admins;

                    // See if we are done
                    if (++ctr >= 2) { resolve(); }
                });
            }

            // Load the owners
            DataSource.loadSiteOwners(this._webUrl).then(owners => {
                // Set the owners
                this._owners = owners;

                // See if we are done
                if (++ctr >= 2) { resolve(); }
            });
        });
    }

    // Shows the remove user dialog
    private removeUser(user: ISiteUserInfo, onRemove: () => void) {
        // Clear the dialog
        Modal.clear();

        // Set the header
        Modal.setHeader("Remove User");

        // Set the content
        Modal.setBody(`Are you sure you want to remove ${user.title} as a ${user.permission} from the site?`);

        // Set the footer
        Components.TooltipGroup({
            el: Modal.FooterElement,
            tooltips: [
                {
                    content: "Click to remove the user from the group.",
                    btnProps: {
                        text: "Remove",
                        onClick: () => {
                            let web = Web(this._webUrl, { requestDigest: DataSource.SiteContext.FormDigestValue });

                            // See if this is an admin or owner
                            if (user.permission === "Admin") {
                                // Get the user
                                web.SiteUsers().getById(user.id).update({ IsSiteAdmin: false }).execute(() => {
                                    // Parse the admins
                                    for (let i = 0; i < this._admins.length; i++) {
                                        let admin = this._admins[i];
                                        if (admin.email === user.email) {
                                            this._admins.splice(i, 1);
                                            break;
                                        }
                                    }

                                    // Call the event
                                    onRemove();
                                });
                                /*
                                // Get the item to update
                                web.Lists().query({
                                    Filter: "BaseTemplate eq " + SPTypes.ListTemplateType.UserInformation
                                }).execute(lists => {
                                    // Get the item
                                    web.Lists().getById(lists.results[0].Id).Items().query({
                                        Filter: "Email eq '" + user.email + "'"
                                    }).execute(items => {
                                        // Update the user
                                        web.Lists().getById(lists.results[0].Id).Items(items.results[0].Id).update({
                                            IsSiteAdmin: false
                                        }).execute(() => {
                                            // Parse the admins
                                            for (let i = 0; i < this._admins.length; i++) {
                                                let admin = this._admins[i];
                                                if (admin.email === user.email) {
                                                    this._admins.splice(i, 1);
                                                    break;
                                                }
                                            }

                                            // Call the event
                                            onRemove();
                                        });
                                    });
                                });
                                */
                            } else {
                                // Get the owners group
                                web.AssociatedOwnerGroup().Users().removeById(user.id).execute(() => {
                                    // Parse the owners
                                    for (let i = 0; i < this._owners.length; i++) {
                                        let owner = this._owners[i];
                                        if (owner.email === user.email) {
                                            this._owners.splice(i, 1);
                                            break;
                                        }
                                    }

                                    // Call the event
                                    onRemove();
                                });
                            }
                        }
                    }
                },
                {
                    content: "Click to close the form.",
                    btnProps: {
                        text: "Close",
                        onClick: () => {
                            Modal.hide();
                        }
                    }
                }
            ]
        });

        // Show the modal
        Modal.show();
    }

    // Renders the audit form
    private render() {
        // Render an alert
        Components.Alert({
            el: this._el,
            header: "Loading User Information",
            content: "Loading the site admins and owners. This will take a few seconds to complete..."
        });

        // Load the users
        this.loadUsers().then(() => {
            // Clear the alert
            while (this._el.firstChild) { this._el.removeChild(this._el.firstChild); }

            // Render the table
            let dt = new Dashboard({
                el: this._el,
                navigation: {
                    items: [
                        {
                            className: "btn-outline-light ms-2",
                            isButton: true,
                            text: "Add User",
                            onClick: () => {
                                // Show the add user dialog
                                this.addUser(() => {
                                    // Refresh the table
                                    dt.refresh(this._admins.concat(this._owners));

                                    // Hide the modal
                                    Modal.hide();
                                });
                            }
                        },
                        {
                            className: "btn-outline-light ms-2",
                            isButton: true,
                            text: "View Permissions",
                            onClick: () => {
                                // Open the site permissions in a new tab
                                window.open(this._webUrl + "/_layouts/15/user.aspx", "_blank");
                            }
                        }
                    ]
                },
                filters: {
                    items: [
                        {
                            header: "By Permission",
                            items: [
                                { label: "Admin", type: Components.CheckboxGroupTypes.Switch },
                                { label: "Owner", type: Components.CheckboxGroupTypes.Switch }
                            ],
                            onFilter: (value: string) => {
                                // Filter the table
                                dt.filter(0, value);
                            }
                        }
                    ]
                },
                table: {
                    onRendering: (dtProps) => {
                        // Disable ordering/searching for the last column
                        dtProps.columnDefs = [
                            {
                                "targets": 3,
                                "orderable": false,
                                "searchable": false
                            }
                        ];

                        // Sort by the first column
                        dtProps.order = [[0, "asc"]];

                        // Return the properties
                        return dtProps;
                    },
                    rows: this._admins.concat(this._owners),
                    columns: [
                        {
                            name: "permission",
                            title: "Permission"
                        },
                        {
                            name: "title",
                            title: "Full Name"
                        },
                        {
                            name: "email",
                            title: "Email"
                        },
                        {
                            name: "",
                            title: "Actions",
                            onRenderCell: (el, row, item: ISiteUserInfo) => {
                                // See if this is an owner and this is an admin item
                                if (!DataSource.IsAdmin && item.permission === "Admin") { return; }

                                // Render the actions
                                Components.TooltipGroup({
                                    el,
                                    isSmall: true,
                                    tooltips: [
                                        {
                                            content: "Click to remove the user from the group.",
                                            btnProps: {
                                                text: "Remove",
                                                onClick: () => {
                                                    // Show the remove user dialog
                                                    this.removeUser(item, () => {
                                                        // Refresh the datatable
                                                        dt.refresh(this._admins.concat(this._owners));

                                                        // Hide the modal
                                                        Modal.hide();
                                                    });
                                                }
                                            }
                                        }
                                    ]
                                })
                            }
                        },
                    ]
                }
            });
        });
    }
}