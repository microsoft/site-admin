import { Dashboard, LoadingDialog, Modal } from "dattatable";
import { Components, Helper, v2 } from "gd-sprest-bs";
import { DataSource } from "../ds";

interface ISitePermission {
    appId: string;
    displayName: string;
    id: string;
    role: string;
}

/**
 * Application Permissions Tab
 */
export class AppPermissionsTab {
    private _el: HTMLElement = null;

    // Constructor
    constructor(el: HTMLElement) {
        this._el = el;

        // Render the tab
        this.render();
    }

    // Adds a permission
    private addPermission(id: string, displayName: string, permission: string): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Show a loading dialog
            LoadingDialog.setHeader("Adding Permission");
            LoadingDialog.setBody("This will close after the permission is added...");
            LoadingDialog.show();

            // Remove the permission
            v2.sites({ siteId: DataSource.Site.Id }).permissions().add({
                grantedToIdentities: [{
                    application: {
                        id,
                        displayName,
                    }
                }],
                roles: [permission]
            }).execute(
                // Success
                () => {
                    // Render the solution
                    this.render();

                    // Hide the loading dialog
                    LoadingDialog.hide();

                    // Resolve the request
                    resolve();
                },

                (err) => {
                    // Reject the request
                    reject(err);

                    // Hide the loading dialog
                    LoadingDialog.hide();
                }
            );
        });
    }

    // Loads the permissions for the site
    private loadPermissions(): PromiseLike<ISitePermission[]> {
        // Return a promise
        return new Promise((resolve, reject) => {
            let site = v2.sites({ siteId: DataSource.Site.Id });
            let sitePermissions: ISitePermission[] = [];

            // Load the permissions
            site.permissions().execute(permissions => {
                // Parse the results
                Helper.Executor(permissions.results, permission => {
                    // Return a promise
                    return new Promise(resolve => {
                        // Get the permission
                        site.permissions(permission.id).execute(permission => {
                            // Append the permission
                            sitePermissions.push({
                                appId: permission.grantedToIdentities[0].application.id,
                                displayName: permission.grantedToIdentities[0].application.displayName,
                                id: permission.id,
                                role: permission.roles[0]
                            });

                            // Resolve the request
                            resolve(null);
                        }, resolve);
                    });
                }).then(() => {
                    // Resolve the request
                    resolve(sitePermissions);
                });
            }, reject);
        });
    }

    // Removes the permission
    private removePermission(permission: ISitePermission) {
        // Show a loading dialog
        LoadingDialog.setHeader("Removing Permission");
        LoadingDialog.setBody("This will close after the permission is removed...");
        LoadingDialog.show();

        // Remove the permission
        v2.sites({ siteId: DataSource.Site.Id }).permissions(permission.id).delete().execute(
            // Success
            () => {
                // Render the solution
                this.render();

                // Hide the loading dialog
                LoadingDialog.hide();
            },

            () => {
                // Hide the loading dialog
                LoadingDialog.hide();
            }
        );
    }

    // Renders the tab
    private render() {
        // Clear the element
        while (this._el.firstChild) { this._el.removeChild(this._el.firstChild); }

        // Render an alert
        Components.Alert({
            el: this._el,
            content: "Loading the app permissions"
        });

        // Load the app permissions
        this.loadPermissions().then(permissions => {
            // Clear the element
            while (this._el.firstChild) { this._el.removeChild(this._el.firstChild); }

            // Render a dashboard
            new Dashboard({
                el: this._el,
                navigation: {
                    showFilter: false,
                    items: [{
                        className: "btn-outline-light",
                        text: "Add",
                        isButton: true,
                        onClick: () => {
                            // Show the new form
                            this.showForm();
                        }
                    }]
                },
                table: {
                    rows: permissions,
                    columns: [
                        {
                            name: "appId",
                            title: "App ID"
                        },
                        {
                            name: "displayName",
                            title: "App Name"
                        },
                        {
                            name: "role",
                            title: "Role"
                        },
                        {
                            name: "",
                            title: "",
                            onRenderCell: (el, col, permission: ISitePermission) => {
                                // Render a button
                                Components.ButtonGroup({
                                    el,
                                    buttons: [
                                        {
                                            text: "Remove",
                                            type: Components.ButtonTypes.OutlineDanger,
                                            onClick: () => {
                                                // Remove the permission
                                                this.removePermission(permission);
                                            }
                                        },
                                        {
                                            text: "Update",
                                            type: Components.ButtonTypes.OutlinePrimary,
                                            onClick: () => {
                                                // Show the form
                                                this.showForm(permission);
                                            }
                                        }
                                    ]
                                });
                            }
                        }
                    ]
                }
            });
        });
    }

    // Displays the new form
    private showForm(permission?: ISitePermission) {
        // Clear the form
        Modal.clear();

        // Set the header
        Modal.setHeader("Add App Permission");

        // Set the form
        let form = Components.Form({
            el: Modal.BodyElement,
            controls: [
                {
                    name: "id",
                    label: "App ID",
                    isDisabled: permission ? true : false,
                    type: Components.FormControlTypes.TextField,
                    required: true,
                    errorMessage: "The app id is required.",
                    value: permission?.appId
                },
                {
                    name: "displayName",
                    label: "App Name",
                    type: Components.FormControlTypes.TextField,
                    required: true,
                    errorMessage: "The app name is required.",
                    value: permission?.displayName
                },
                {
                    name: "permission",
                    label: "Permission:",
                    type: Components.FormControlTypes.Dropdown,
                    required: true,
                    value: permission?.role,
                    items: [
                        { text: "Read", value: "read" },
                        { text: "Write", value: "write" },
                        { text: "Full Control", value: "fullcontrol" }
                    ]
                } as Components.IFormControlPropsDropdown
            ]
        });

        // Set the footer
        Components.Button({
            el: Modal.FooterElement,
            text: permission ? "Update" : "Add",
            type: Components.ButtonTypes.OutlinePrimary,
            onClick: () => {
                // Ensure it's valid
                if (form.isValid()) {
                    let values = form.getValues();

                    // See if the permission exists
                    if (permission) {
                        // Update the permission
                        permission.role = values.permission.value;
                        this.updatePermission(permission).then(() => {
                            // Hide the modal
                            Modal.hide();
                        }, () => {
                            // Show an error
                            let ctrl = form.getControl("id");
                            ctrl.updateValidation(ctrl.el, {
                                isValid: false,
                                invalidMessage: "There was an error updating the permission."
                            });
                        });
                    } else {
                        // Add the permission
                        this.addPermission(values.id, values.displayName, values.permission.value).then(() => {
                            // Hide the modal
                            Modal.hide();
                        }, () => {
                            // Show an error
                            let ctrl = form.getControl("id");
                            ctrl.updateValidation(ctrl.el, {
                                isValid: false,
                                invalidMessage: "There was an error adding the permission."
                            });
                        });
                    }
                }
            }
        });

        // Show the form
        Modal.show();
    }

    // Adds a permission
    private updatePermission(permission: ISitePermission): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Show a loading dialog
            LoadingDialog.setHeader("Updating Permission");
            LoadingDialog.setBody("This will close after the permission is updated...");
            LoadingDialog.show();

            // Remove the permission
            v2.sites({ siteId: DataSource.Site.Id }).permissions(permission.id).update({
                roles: [permission.role]
            }).execute(
                // Success
                () => {
                    // Render the solution
                    this.render();

                    // Hide the loading dialog
                    LoadingDialog.hide();

                    // Resolve the request
                    resolve();
                },

                (err) => {
                    // Reject the request
                    reject(err);

                    // Hide the loading dialog
                    LoadingDialog.hide();
                }
            );
        });
    }
}