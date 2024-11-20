import { LoadingDialog } from "dattatable";
import { Components, Types } from "gd-sprest-bs";
import { DataSource, RequestTypes, IAPIRequestProps } from "../ds";
import { APIResponseModal } from "./response";

/**
 * Site Form
 */
export class Site {
    private _apiUrls: string[] = null;
    private _el: HTMLElement = null;
    private _form: Components.IForm = null;
    private _disableProps: string[] = null;
    private _site: Types.SP.SiteOData = null;

    // The current values
    private _currValues: {
        CommentsOnSitePagesDisabled: boolean;
        ContainsAppCatalog: boolean;
        DisableCompanyWideSharingLinks: boolean;
        ShareByEmailEnabled: boolean;
        SocialBarOnSitePagesDisabled: boolean;
    } = null;

    constructor(site: Types.SP.SiteOData, el: HTMLElement, disableProps: string[] = [], apiUrls: string[] = []) {
        // Save the properties
        this._apiUrls = apiUrls;
        this._el = el;
        this._disableProps = disableProps;
        this._site = site;

        // Set the current values
        this._currValues = {
            CommentsOnSitePagesDisabled: this._site.CommentsOnSitePagesDisabled,
            ContainsAppCatalog: false,
            DisableCompanyWideSharingLinks: this._site.DisableCompanyWideSharingLinks,
            ShareByEmailEnabled: this._site.ShareByEmailEnabled,
            SocialBarOnSitePagesDisabled: this._site.SocialBarOnSitePagesDisabled
        }

        // Clear the element
        while (el.firstChild) { el.removeChild(el.firstChild); }

        // Render the header
        this.renderHeader();

        // Renders the form
        this.renderForm();

        // Renders the footer
        this.renderFooter();
    }

    // Renders the footer
    private renderFooter() {
        // Create the element
        let elFooter = document.createElement("div");
        elFooter.classList.add("mt-5");
        elFooter.classList.add("d-flex");
        elFooter.classList.add("justify-content-end");
        this._el.appendChild(elFooter);

        // Render the tooltips
        Components.Tooltip({
            el: elFooter,
            content: "Click to save the changes.",
            btnProps: {
                text: "Save Changes",
                onClick: () => {
                    let values = this._form.getValues();

                    // Save the site properties
                    this.save(values);
                }
            }
        });
    }

    // Renders the form
    private renderForm() {
        // Render the form
        this._form = Components.Form({
            el: this._el,
            groupClassName: "mb-3",
            controls: [
                {
                    name: "CommentsOnSitePagesDisabled",
                    label: "Comments On Site Pages Disabled:",
                    description: "The type of web.",
                    isDisabled: this._disableProps.indexOf("CommentsOnSitePagesDisabled") >= 0,
                    type: Components.FormControlTypes.Switch,
                    value: this._currValues.CommentsOnSitePagesDisabled
                } as Components.IFormControlPropsSwitch,
                {
                    name: "DisableCompanyWideSharingLinks",
                    label: "Disable Company Wide Sharing Links:",
                    description: "If true, it will hide the comments on the site pages.",
                    isDisabled: this._disableProps.indexOf("DisableCompanyWideSharingLinks") >= 0,
                    type: Components.FormControlTypes.Switch,
                    value: this._currValues.DisableCompanyWideSharingLinks
                } as Components.IFormControlPropsSwitch,
                {
                    name: "ContainsAppCatalog",
                    label: "App Catalog Enabled:",
                    description: "True if this has a site collection app catalog available.",
                    isDisabled: this._disableProps.indexOf("ContainsAppCatalog") >= 0,
                    type: Components.FormControlTypes.Switch,
                    onControlRendering: ctrl => {
                        // Return a promise
                        return new Promise((resolve) => {
                            // See if an app catalog exists on this site
                            DataSource.hasAppCatalog(this._site.Url).then(hasAppCatalog => {
                                // Set the value
                                this._currValues.ContainsAppCatalog = hasAppCatalog;
                                ctrl.value = hasAppCatalog

                                // Resolve the request
                                resolve(ctrl);
                            });
                        });
                    }
                } as Components.IFormControlPropsSwitch,
                {
                    name: "ShareByEmailEnabled",
                    label: "Share By Email Enabled:",
                    description: "Disables the offline sync feature in all libraries.",
                    isDisabled: this._disableProps.indexOf("ShareByEmailEnabled") >= 0,
                    type: Components.FormControlTypes.Switch,
                    value: this._currValues.ShareByEmailEnabled
                } as Components.IFormControlPropsSwitch,
                {
                    name: "SocialBarOnSitePagesDisabled",
                    label: "Social Bar On Site Pages Disabled:",
                    description: "The search scope for the site to target. Default is 'Site'.",
                    isDisabled: this._disableProps.indexOf("SocialBarOnSitePagesDisabled") >= 0,
                    type: Components.FormControlTypes.Switch,
                    value: this._currValues.SocialBarOnSitePagesDisabled
                } as Components.IFormControlPropsSwitch
            ]
        });
    }

    // Renders the header
    private renderHeader() {
        // Render the header
        Components.Jumbotron({
            el: this._el,
            className: "mb-2",
            lead: "Site Collection Settings",
            size: Components.JumbotronSize.Small,
            type: Components.JumbotronTypes.Primary
        });
    }

    // Saves the properties
    private save(values): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve) => {
            let apis: IAPIRequestProps[] = [];
            let props = {};
            let requests: string[] = [];
            let updateFl = false;

            // Parse the keys
            for (let key in this._currValues) {
                let value = typeof (values[key]) === "boolean" ? values[key] : values[key].data;
                if (this._currValues[key] != value) {
                    // See if there is an associated api
                    let keyIdx = this._apiUrls.indexOf(key);
                    if (keyIdx >= 0) {
                        // Append the url
                        apis.push({ key, value, api: this._apiUrls[keyIdx] });
                    }
                    // Else, see if we need to create a request
                    else if (key == "ContainsAppCatalog") {
                        // Add a request for this request
                        requests.push(RequestTypes.AppCatalog);
                    }
                    // Else, see if we need to create a request
                    else if (key == "DisableCompanyWideSharingLinks") {
                        // Add a request for this request
                        requests.push(RequestTypes.DisableCompanyWideSharingLinks);
                    }
                    // Else, we can update this using REST
                    else {
                        // Add the property
                        props[key] = value;

                        // Set the flag
                        updateFl = true;
                    }
                }
            }

            // Add the requests
            DataSource.addRequest(this._site.Url, requests).then((responses) => {
                // Process the requests
                DataSource.processAPIRequests(apis).then((apiResponses) => {
                    // See if an update is needed
                    if (updateFl) {
                        // Show a loading dialog
                        LoadingDialog.setHeader("Updating Site Collection");
                        LoadingDialog.setBody("This will close after the changes complete.");
                        LoadingDialog.show();

                        // Save the changes
                        this._site.update(props).execute(() => {
                            // Update the current values
                            this._currValues.CommentsOnSitePagesDisabled = values["CommentsOnSitePagesDisabled"];
                            this._currValues.ShareByEmailEnabled = values["ShareByEmailEnabled"];
                            this._currValues.SocialBarOnSitePagesDisabled = values["SocialBarOnSitePagesDisabled"];

                            // Close the dialog
                            LoadingDialog.hide();

                            // Show the responses
                            new APIResponseModal(responses.concat(apiResponses));

                            // Resolve the request
                            resolve();
                        });
                    } else {
                        // Show the responses
                        new APIResponseModal(responses.concat(apiResponses));

                        // Resolve the request
                        resolve();
                    }
                });

            });
        });
    }
}