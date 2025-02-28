import { LoadingDialog, Modal } from "dattatable";
import { Components, Search } from "gd-sprest-bs";
import { DataSource } from "./ds";

/**
 * Load Form
 */
export class LoadForm {
    _form: Components.IForm = null;
    _onSuccess: () => void = null;

    // Renders the modal
    constructor(elForm: HTMLElement, elFooter: HTMLElement, onSuccess: () => void) {
        this._onSuccess = onSuccess;

        // Render the form and footer
        this.renderForm(elForm);
        this.renderFooter(elFooter);
    }

    // Renders the footer
    private renderFooter(el: HTMLElement) {
        // Clear the element
        while (el.firstChild) { el.removeChild(el.firstChild); }

        // Render the footer
        Components.TooltipGroup({
            el,
            tooltips: [
                {
                    content: "Validates that you are an admin for the site entered.",
                    btnProps: {
                        text: "Load Site",
                        onClick: () => {
                            // Submit the request
                            this.submitForm();
                        }
                    }
                }
            ]
        });
    }

    // Renders the form
    private renderForm(el: HTMLElement) {
        let disableEvent = false;
        let popover: Components.IPopover = null;
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
                    placeholder: "Absolute or Relative Site Url",
                    type: Components.FormControlTypes.TextField,
                    description: "The absolute/relative url to the site. (Example: /sites/dev)<br/>Type in a minimum of 3 characters to search for sites.",
                    required: true,
                    errorMessage: "The site url is required.",
                    prependedDropdown: {
                        updateLabel: true,
                        items: [
                            { text: "/sites/", value: "/sites/", isSelected: true },
                            { text: "/teams/", value: "/teams/" }
                        ]
                    },
                    onControlRendered: ctrl => {
                        // Set the key down event
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
                        popover = Components.Popover({
                            className: "search-sites",
                            target: tb.elTextbox,
                            placement: Components.PopoverPlacements.BottomStart,
                            options: {
                                showOnCreate: false,
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
                            popover.setBody("Searching for sites...");
                            popover.show();

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
                                popover.setBody(Components.Dropdown({
                                    menuOnly: true,
                                    items,
                                    onChange: ((item: Components.IDropdownItem) => {
                                        // Set the value
                                        disableEvent = true;
                                        tb.setValue(item?.value);
                                        disableEvent = false;

                                        // Hide the popover
                                        popover.hide();
                                    })
                                }).el);
                            }, () => {
                                // Error searching for sites
                                popover.hide();
                            });
                        }, 100);
                    }
                } as Components.IFormControlPropsTextField
            ]
        });
    }

    // Shows the modal
    static showModal(onSuccess: () => void) {
        // Clear the modal
        Modal.clear();

        // Set the header
        Modal.setHeader("Load Site");

        // Create the form/footer
        new LoadForm(Modal.BodyElement, Modal.FooterElement, onSuccess);

        // Show the modal
        Modal.show();
    }

    // Submits the form
    private submitForm() {
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
}