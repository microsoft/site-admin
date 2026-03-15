import { LoadingDialog, Modal, Navigation } from "dattatable";
import { Components } from "gd-sprest-bs";
import { generateIcon } from "gd-sprest-bs/build/icons/generate.js";
import { DataSource, RequestTypes } from "./ds";
import { InstallationModal } from "./install";
import { LoadForm } from "./loadForm";
import Strings from "./strings";
import { Security } from "./security";
import { SiteAttestationForm } from "./siteAttestationForm";
import { Tabs } from "./tabs";
import { IReportProps } from "./tabs/reports";
import { ISearchProps } from "./tabs/searchProp";

// App Properties
export interface IProp {
    description: string;
    disabled: boolean;
    label: string;
}
export interface IAppProps {
    auditOnly?: boolean;
    context?: any;
    disableSensitivityLabelOverride?: boolean;
    el: HTMLElement;
    hideCreateSiteBtn?: boolean;
    hideLoadAdminOwnerBtn?: boolean;
    hideLoadOneDriveBtn?: boolean;
    hideReports: {
        dlp?: boolean;
        docRetention?: boolean;
        externalShares?: boolean;
        externalUsers?: boolean;
        permissions?: boolean;
        retention?: boolean;
        searchDocs?: boolean;
        searchEEEU?: boolean;
        searchProp?: boolean;
        searchUsers?: boolean;
        sensitivityLabels?: boolean;
        sharingLinks?: boolean;
        uniquePermissions?: boolean;
    }
    hideTabs: {
        appPermissions?: boolean;
        auditTools?: boolean;
        features?: boolean;
        lists?: boolean;
        management?: boolean;
        search?: boolean;
        webs?: boolean;
    }
    imageReferences: string[];
    maxRequests?: number;
    maxStorageDesc?: string;
    maxStorageSize?: number;
    reportProps?: IReportProps;
    searchProps?: ISearchProps;
    siteAttestation?: boolean;
    siteAttestationText?: string;
    siteProps: { [key: string]: IProp; }
    title?: string;
    webProps: { [key: string]: IProp; }
}

/**
 * Main Application
 */
export class App {
    private _props: IAppProps = null;

    // Constructor
    constructor(props: IAppProps, loadOneDrive: boolean) {
        this._props = props;

        // Add the class for bootstrap
        this._props.el.classList.add("bs");

        // Render the template
        this._props.el.innerHTML = `
            <div class="row">
                <div class="col-12"></div>
                <div class="col-12 mt-2"></div>
            </div>
        `;

        // Set the elements
        let elRow = this._props.el.children[0] as HTMLElement;

        // Render the dashboard
        this.renderNavigation(elRow, loadOneDrive);

        // See if we are loading onedrive
        if (loadOneDrive) {
            // Render the tabs
            this.renderTabs(elRow.children[1] as HTMLElement, true);
        }
        // Else, see if data has been loaded
        else if (DataSource.Site) {
            // Render the tabs
            this.renderTabs(elRow.children[1] as HTMLElement, false);
        } else {
            // Render the load form
            this.renderForm(elRow.children[1] as HTMLElement);
        }
    }

    // Renders the form
    private renderForm(el: HTMLElement) {
        // Render the form
        el.innerHTML = `
            <div class="row">
                <div class="col-12 my-3"></div>
                <div class="col-12 d-flex justify-content-end"></div>
            </div>
        `;

        // Render the form
        new LoadForm(el.children[0].children[0] as HTMLElement, el.children[0].children[1] as HTMLElement,
            this._props.hideLoadAdminOwnerBtn, this._props.hideLoadOneDriveBtn, () => {
                // Render the tabs
                new App(this._props, false);
            }, () => {
                // Render the tabs
                new App(this._props, true);
            });
    }

    // Renders the navigation
    private renderNavigation(elRow: HTMLElement, loadOneDrive: boolean) {
        let itemsEnd: Components.INavbarItem[] = [];

        // See if we are showing the create site button
        if (!this._props.hideCreateSiteBtn) {
            // Add the create site button
            itemsEnd.push({
                className: "btn-outline-light ms-2",
                isButton: true,
                text: "Create Site",
                onClick: () => {
                    // Show the create site form
                    this.showCreateSiteForm();
                }
            })
        }

        // Show the load site button if data has already been loaded
        if (DataSource.Site || loadOneDrive) {
            itemsEnd.push({
                className: "btn-outline-light ms-2",
                isButton: true,
                text: "Load Site",
                onClick: () => {
                    // Show the load form
                    LoadForm.showModal(this._props.hideLoadAdminOwnerBtn, this._props.hideLoadOneDriveBtn, () => {
                        // Render the tabs
                        this.renderTabs(elRow.children[1] as HTMLElement, false);
                    }, () => {
                        // Render the tabs
                        this.renderTabs(elRow.children[1] as HTMLElement, true);
                    });
                }
            });

            // See if we are enabling the site attestation feature
            if (this._props.siteAttestation) {
                // Add the settings for the app
                itemsEnd.push({
                    className: "btn-outline-light ms-2",
                    isButton: true,
                    text: "Site Attestation",
                    onClick: () => {
                        // Show the site attestation form
                        new SiteAttestationForm(this._props.siteAttestationText);
                    }
                });
            }
        }

        // See if this is the admin
        if (Security.IsAdmin) {
            // Add the settings for the app
            itemsEnd.push({
                className: "btn-outline-light ms-2",
                isButton: true,
                text: "Settings",
                onClick: () => {
                    // Show the app settings
                    InstallationModal.show(true);
                }
            });
        }

        // Render the navigation
        new Navigation({
            iconType: generateIcon(`<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 2048 2048"><path fill="var(--bs-navbar-brand-color)" d="M64 272L1216 67v1914L64 1782V272zm613 347q-54 0-98 17t-75 50-48 76-17 98q0 59 18 98t49 69 68 55 78 57q17 14 26 33t9 42q0 38-23 58t-60 20q-23 0-46-6t-44-18-40-28-32-35v173q39 26 84 40t93 14q50 0 91-15t70-45 44-70 16-92q0-62-19-104t-49-72-62-50-63-41-48-43-20-58q0-22 8-38t23-26 33-14 39-5q33 0 69 13t60 38V641q-31-11-66-16t-68-6zm731 405q0 26-10 49t-27 41-41 28-50 10V897q27 0 50 10t40 27 28 40 10 50zm616 86l-63 241-212-28q-17 26-35 50t-41 46l74 192-216 126-128-168q-60 15-123 15v-210q72 0 135-27t111-75 74-110 28-136q0-72-27-135t-75-111-110-75-136-28V461h5l79-180 241 64-27 205q27 17 52 37t48 44l188-72 125 215-167 129q13 59 13 124l187 83z"></path></svg>`, 32, 32),
            el: elRow.children[0] as HTMLElement,
            title: this._props.title || Strings.ProjectName,
            hideFilter: true,
            hideSearch: true,
            itemsEnd
        });
    }

    // Renders the tabs
    private renderTabs(el: HTMLElement, loadOneDrive: boolean) {
        // Clear the tabs element
        while (el.firstChild) { el.removeChild(el.firstChild); }

        // Render the site information
        Components.Form({
            el,
            className: "mt-1 mb-4",
            rows: [
                {
                    columns: loadOneDrive ? [
                        {
                            control: {
                                label: "OneDrive Site Url:",
                                type: Components.FormControlTypes.Readonly,
                                value: DataSource.OneDriveWeb.Url
                            }
                        }
                    ] : [
                        {
                            size: 6,
                            control: {
                                label: "Top Site Url:",
                                type: Components.FormControlTypes.Readonly,
                                value: DataSource.Site.Url
                            }
                        },
                        {
                            size: 6,
                            control: {
                                label: "Sub Site Url:",
                                type: Components.FormControlTypes.Dropdown,
                                items: [{ text: "Loading Sites..." }],
                                value: DataSource.Web.Id,
                                required: true,
                                onChange: item => {
                                    // Ensure it's a valid site
                                    if (item.value) {
                                        // Refresh the web tab
                                        tabs.refreshWebTab(item.text);
                                    }
                                },
                                onControlRendered: ctrl => {
                                    // Load the sub-webs
                                    DataSource.getAllWebs(DataSource.Site.Url).then(() => {
                                        // Update the tabs
                                        tabs.onWebsLoaded(DataSource.SiteItems);

                                        // Update the control
                                        ctrl.dropdown.setItems(DataSource.SiteItems);
                                    });
                                }
                            } as Components.IFormControlPropsDropdown
                        }
                    ]
                }
            ]
        });

        // Render the tabs
        let tabs = new Tabs(el, this._props, loadOneDrive);
    }

    // The create site form
    private showCreateSiteForm() {
        // Clear the modal
        Modal.clear();

        // Set the header
        Modal.setHeader("Create Site");

        // Set the form
        let form = Components.Form({
            el: Modal.BodyElement,
            controls: [
                {
                    name: "Title",
                    label: "Site Title:",
                    type: Components.FormControlTypes.TextField,
                    required: true,
                    errorMessage: "The name of the site is required.",
                    onChange: value => {
                        // Set the default url value
                        form.getControl("GroupAlias").textbox.setValue(value.replace(/\s/g, ""));
                        form.getControl("Url").textbox.setValue(value.replace(/\s/g, "").toLowerCase());
                    }
                } as Components.IFormControlPropsTextField,
                {
                    name: "GroupAlias",
                    label: "Group Alias:",
                    type: Components.FormControlTypes.TextField,
                    onValidate: (ctrl, results) => {
                        // See what template is selected
                        let template = form.getControl("Template").getValue().value;
                        if (template != "STS#3") {
                            // Ensure a value exists
                            results.isValid = results.value ? true : false;
                            results.invalidMessage = "A group alias is required for M365 Group connected sites.";
                        }

                        // Return the results
                        return results;
                    }
                } as Components.IFormControlPropsTextField,
                {
                    name: "Url",
                    label: "Site URL:",
                    type: Components.FormControlTypes.TextField,
                    required: true,
                    errorMessage: "The URL of the site is required.",
                    prependedDropdown: {
                        updateLabel: true,
                        items: [
                            { text: "/sites/", value: "/sites/", isSelected: true },
                            { text: "/teams/", value: "/teams/" }
                        ]
                    }
                } as Components.IFormControlPropsTextField,
                {
                    name: "Description",
                    label: "Site Description:",
                    type: Components.FormControlTypes.TextField
                } as Components.IFormControlPropsTextField,
                {
                    name: "Template",
                    label: "Site Template:",
                    type: Components.FormControlTypes.Dropdown,
                    required: true,
                    errorMessage: "A site template is required.",
                    items: [
                        { text: "Team Site (No M365 Group)", value: "STS#3" },
                        { text: "Team Site (M365 Group)", value: "GROUP#0" },
                        { text: "Communication Site", value: "SITEPAGEPUBLISHING#0" }
                    ]
                } as Components.IFormControlPropsDropdown
            ]
        });

        // Set the footer
        Components.TooltipGroup({
            el: Modal.FooterElement,
            tooltips: [
                {
                    content: "Submits the request to create a new site.",
                    btnProps: {
                        text: "Create",
                        type: Components.ButtonTypes.OutlinePrimary,
                        onClick: () => {
                            // Ensure the form is valid
                            if (form.isValid()) {
                                let values = form.getValues();
                                let validateGroupAlias = values["Template"].value == "STS#3";

                                // Show a loading dialog
                                LoadingDialog.setHeader("Validating the Site Information");
                                LoadingDialog.setBody("Validating the site information...");
                                LoadingDialog.show();

                                // Validate the site info
                                DataSource.validateSiteCreationInfo(values["Title"], validateGroupAlias ? values["GroupAlias"] : null, values["Url"]).then(() => {
                                    // Update the dialog
                                    LoadingDialog.setHeader("Creating Site Request");
                                    LoadingDialog.setBody("Creating the request for a new site. This will close when it completes.");

                                    // Create the request for the site
                                    DataSource.addRequest(values["Url"], [{
                                        key: RequestTypes.CreateSite,
                                        message: "",
                                        value: JSON.stringify({
                                            Description: values["Description"],
                                            GroupAlias: values["GroupAlias"],
                                            Template: values["Template"].value,
                                            Title: values["Title"],
                                            Url: values["Url"]
                                        })
                                    }]).then(() => {
                                        // Hide the dialogs
                                        Modal.hide();
                                        LoadingDialog.hide();
                                    });
                                }, (invalidMessage: string) => {
                                    // Update the validation
                                    let ctrl = invalidMessage.startsWith("A site already exists") ? form.getControl("Url") : form.getControl("GroupAlias");

                                    // Update the validation
                                    ctrl.updateValidation(ctrl.el, {
                                        isValid: false,
                                        invalidMessage
                                    });

                                    // Hide the dialog
                                    LoadingDialog.hide();
                                    return;
                                });
                            }
                        }
                    }
                },
                {
                    content: "Closes the form.",
                    btnProps: {
                        text: "Close",
                        type: Components.ButtonTypes.OutlineSecondary,
                        onClick: () => {
                            // Hide the modal
                            Modal.hide();
                        }
                    }
                }
            ]
        });

        // Show the form
        Modal.show();
    }
}