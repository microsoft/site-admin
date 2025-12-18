import { Dashboard, LoadingDialog, Modal, DataTable } from "dattatable";
import { Components, Search } from "gd-sprest-bs";
import { DataSource } from "./ds";
import { PageGenerator } from "./page-generator";
import { Security } from "./security";

/**
 * Load Form
 */
export class LoadForm {
    private _form: Components.IForm = null;
    private _popover: Components.IPopover = null;
    private _onSuccess: () => void = null;

    // Renders the modal
    constructor(elForm: HTMLElement, elFooter: HTMLElement, hideLoadAdminOwnerBtn: boolean = false, onSuccess: () => void) {
        this._onSuccess = onSuccess;

        // Render the form and footer
        this.renderForm(elForm);
        this.renderFooter(elFooter, hideLoadAdminOwnerBtn);
    }

    // Renders the footer
    private renderFooter(el: HTMLElement, hideLoadAdminOwnerBtn: boolean) {
        // Clear the element
        while (el.firstChild) { el.removeChild(el.firstChild); }

        // Set the tooltips
        let tooltips = []
        if (!hideLoadAdminOwnerBtn) {
            tooltips.push({
                content: "Views the owners/admins of the site.",
                btnProps: {
                    text: "View Admins/Owners",
                    onClick: () => {
                        // Load the admins/owners of the site
                        this.viewAdminsOwners();
                    }
                }
            });
        }
        tooltips.push({
            content: "Validates that you are an admin for the site entered.",
            btnProps: {
                text: "Load Site",
                onClick: () => {
                    // Submit the request
                    this.submitForm();
                }
            }
        });

        // Render the footer
        Components.TooltipGroup({
            el,
            tooltips
        });
    }

    // Renders the form
    private renderForm(el: HTMLElement) {
        let disableEvent = false;
        let ddl: Components.IDropdown = null;
        let ddlLastSelection = "/sites/";
        let tb: Components.IInputGroup = null;

        // Clear the element
        while (el.firstChild) { el.removeChild(el.firstChild); }

        // Render the form
        this._form = Components.Form({
            el,
            controls: [
                {
                    name: "url",
                    label: "Site Url:",
                    placeholder: "Relative Site Url",
                    type: Components.FormControlTypes.TextField,
                    description: "The relative url to the site. (Example: Enter 'dev' to load: /sites/dev)<br/>Type in a minimum of 3 characters to search for sites.",
                    required: true,
                    errorMessage: "The site url is required.",
                    prependedDropdown: {
                        updateLabel: true,
                        items: [
                            { text: "/", value: "/" },
                            { text: "/sites/", value: "/sites/", isSelected: true },
                            //{ text: "/personal/", value: "/personal/" },
                            { text: "/teams/", value: "/teams/" }
                        ],
                        assignTo: obj => { ddl = obj; },
                        onChange: ((item: Components.IDropdownItem) => {
                            if (item) {
                                // Set the last selection
                                ddlLastSelection = item.value;

                                // See if this is the absolute root
                                if (item.value == "/") {
                                    // Clear and disable the textbox
                                    tb.setValue("");
                                    tb.elTextbox.disabled = true;
                                } else {
                                    // Enable the textbox
                                    tb.elTextbox.disabled = false;
                                }
                            } else {
                                // Set the default item
                                ddl.setValue(ddlLastSelection);
                            }
                        })
                    },
                    onControlRendered: ctrl => {
                        let keyIdx = 0;
                        let keys = [38, 38, 40, 40, 37, 39, 37, 39];

                        // See if the user is an admin
                        if (Security.IsAdmin) {
                            // Set the key down event
                            ctrl.textbox.elTextbox.addEventListener("keydown", ev => {
                                // See if we match the keys
                                if (ev["keyCode"] === keys[keyIdx]) {
                                    // See if we have matched all the keys
                                    if (++keyIdx >= keys.length) {
                                        // Show the page generator form
                                        new PageGenerator();

                                        // Reset the index
                                        keyIdx = 0;
                                    }
                                } else {
                                    // Reset the key index
                                    keyIdx = 0;
                                }
                            });
                        }

                        // Set the key press event
                        ctrl.textbox.elTextbox.addEventListener("keypress", ev => {
                            // See if they hit the enter button
                            if (ev["keyCode"] === 13) {
                                ev.preventDefault();

                                // Submit the request
                                this.submitForm();
                            }
                        });

                        // Create a popover menu
                        tb = ctrl.textbox;
                        this._popover = Components.Popover({
                            className: "search-sites",
                            target: tb.elTextbox,
                            placement: Components.PopoverPlacements.BottomStart,
                            options: {
                                trigger: ""
                            }
                        });
                    },
                    onChange: (value) => {
                        // See if we are disabling the event
                        if (disableEvent) { return; }

                        // Ensure 3 characters exist
                        if (tb.elTextbox.value.length < 3) { return; }

                        // Wait for the user to stop typing
                        let prevValue = value;
                        setTimeout(() => {
                            // See if the value changed
                            if (prevValue != value) { return; }

                            // Show the popover
                            this._popover.setBody("Searching for sites...");
                            this._popover.show();

                            // Determine the path query
                            let selectedItem = tb.prependedDropdown.getValue() as Components.IDropdownItem;
                            let siteUrl = `${document.location.origin}${selectedItem.value}`
                            let pathQuery = `${siteUrl}${tb.elTextbox.value}*`;

                            // Query the search api for sites
                            Search.postQuery<{
                                Path: string;
                                SiteId: string;
                                Title: string;
                            }>({
                                getAllItems: true,
                                query: {
                                    Querytext: `Path:${pathQuery} contentclass=sts_site`,
                                    RowLimit: 15,
                                    TrimDuplicates: true,
                                    SelectProperties: {
                                        results: [
                                            "Path", "SiteId", "Title"
                                        ]
                                    }
                                }
                            }).then(search => {
                                // Parse the results
                                let items: Components.IDropdownItem[] = [];
                                if (search.results.length == 0) {
                                    // Set the default value
                                    items.push({
                                        text: "No sites were found for: " + value,
                                        value: ""
                                    });
                                } else {
                                    // Parse the results
                                    for (let i = 0; i < search.results.length; i++) {
                                        let result = search.results[i];

                                        // Append the site suggestion
                                        items.push({
                                            data: result,
                                            text: result.Path,
                                            value: result.Path.replace(siteUrl, '')
                                        });
                                    }
                                }

                                // Sort the items
                                items.sort((a, b) => {
                                    if (a.text < b.text) { return -1; }
                                    if (a.text > b.text) { return 1; }
                                    return 0;
                                });

                                // Update the popover
                                this._popover.setBody(Components.Dropdown({
                                    menuOnly: true,
                                    items,
                                    onChange: ((item: Components.IDropdownItem) => {
                                        // Set the value
                                        disableEvent = true;
                                        tb.setValue(item?.value);
                                        disableEvent = false;

                                        // Hide the popover
                                        this._popover.hide();
                                    })
                                }).el);
                            }, () => {
                                // Error searching for sites
                                this._popover.hide();
                            });
                        }, 100);
                    }
                } as Components.IFormControlPropsTextField
            ]
        });
    }

    // Shows the modal
    static showModal(hideLoadAdminOwnerBtn: boolean = false, onSuccess: () => void) {
        // Clear the modal
        Modal.clear();

        // Set the header
        Modal.setHeader("Load Site");

        // Create the form/footer
        new LoadForm(Modal.BodyElement, Modal.FooterElement, hideLoadAdminOwnerBtn, onSuccess);

        // Show the modal
        Modal.show();
    }

    // Submits the form
    private submitForm() {
        // Ensure the popover is hidden
        this._popover.hide();

        // Validate the form
        if (this._form.isValid()) {
            let ctrl = this._form.getControl("url");
            let url = this._form.getValues()["url"];

            // Show a loading dialog
            LoadingDialog.setHeader("Validating Site");
            LoadingDialog.setBody("This will close after the site url is validated...");
            LoadingDialog.show();

            // Validate the url
            DataSource.validate(typeof (url) === "string" ? url : url.value).then(
                // Success
                () => {
                    // Close the dialogs
                    LoadingDialog.hide();
                    Modal.hide();

                    // Call the success method
                    this._onSuccess();
                },

                // Error
                (errorMessage) => {
                    // Close the dialog
                    LoadingDialog.hide();

                    // Update the validation
                    ctrl.updateValidation(ctrl.el, {
                        isValid: false,
                        invalidMessage: errorMessage
                    });
                }
            )
        }
    }

    // Views the admins/owners of the site
    private viewAdminsOwners() {
        // Ensure the popover is hidden
        this._popover.hide();

        // Validate the form
        if (this._form.isValid()) {
            let ctrl = this._form.getControl("url");
            let url = this._form.getValues()["url"];

            // Show a loading dialog
            LoadingDialog.setHeader("Validating Site");
            LoadingDialog.setBody("This will close after the site url is validated...");
            LoadingDialog.show();

            // Load the admins/owners
            DataSource.loadAdminsOwners(typeof (url) === "string" ? url : url.value).then(
                // Success
                users => {
                    // Clear the modal
                    Modal.clear();

                    // Set the header
                    Modal.setHeader("Site Admins/Owners");

                    // Set the body
                    new DataTable({
                        el: Modal.BodyElement,
                        rows: users.admins.concat(users.owners),
                        columns: [
                            {
                                name: "type",
                                title: "User Type"
                            },
                            {
                                name: "name",
                                title: "Name"
                            },
                            {
                                name: "email",
                                title: "Email"
                            }
                        ]
                    });

                    // Render the footer
                    Components.TooltipGroup({
                        el: Modal.FooterElement,
                        tooltips: [
                            {
                                content: "Closes the form.",
                                btnProps: {
                                    text: "Close",
                                    onClick: () => {
                                        // Hide the modal
                                        Modal.hide();
                                    }
                                }
                            }
                        ]
                    });

                    // Close the dialog
                    LoadingDialog.hide();

                    // Show the modal
                    Modal.show();
                },

                // Error
                () => {
                    // Close the dialog
                    LoadingDialog.hide();

                    // Update the validation
                    ctrl.updateValidation(ctrl.el, {
                        isValid: false,
                        invalidMessage: "Unable to load the web information. Please check the url: " + url
                    });
                }
            );
        }
    }
}