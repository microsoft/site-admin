import { CanvasForm, Dashboard, Documents, LoadingDialog, Modal } from "dattatable";
import { Components, Helper, SPTypes, Types, Web, v2 } from "gd-sprest-bs";
import { DataSource } from "../ds";
import { M365Groups } from "../m365Groups";
import { BulkLabel } from "./bulkLabel";
import { ExportCSV } from "./exportCSV";
import { ViewPermissions } from "./viewPermissions";

export interface ISensitivityLabelItem {
    Author: string;
    File: Types.Microsoft.Graph.driveItem;
    FileExtension: string;
    FileName: string;
    HasUniquePermissions: boolean;
    ItemId: number;
    ListId: string;
    ListTitle: string;
    Overshared: string;
    Path: string;
    Permissions: Types.SP.RoleAssignmentOData[];
    SensitivityLabel: string;
    SensitivityLabelId: string;
    WebUrl: string;
    WebId: string;
}

interface IWebItem {
    ListId: string;
    ListTitle: string;
    WebId: string;
    WebUrl: string;
}

const CSVFields = [
    "Author",
    "FileExtension",
    "FileName",
    "ListId",
    "ListTitle",
    "Path",
    "SensitivityLabel",
    "SensitivityLabelId",
    "WebId",
    "WebUrl"
]

export class SensitivityLabels {
    private static _dashboard: Dashboard = null;
    private static _elSubNav: HTMLElement = null;
    private static _filterLabels: string[] = [];
    private static _items: ISensitivityLabelItem[] = [];
    private static _loadOneDrive: boolean = false;
    private static _stopFl: boolean = false;

    // Gets the form fields to display
    static getFormFields(): Components.IFormControlProps[] {
        return [
            {
                name: "SearchType",
                className: "my-3",
                type: Components.FormControlTypes.MultiSwitch,
                required: true,
                errorMessage: "A selection is required.",
                items: [
                    {
                        name: "WithLabels",
                        label: "Find all files with a label",
                        isSelected: true
                    },
                    {
                        name: "WithoutLabels",
                        label: "Find all files without a label"
                    }
                ]
            } as Components.IFormControlPropsMultiSwitch,
            {
                name: "FilterLabel",
                className: "mb-3",
                type: Components.FormControlTypes.MultiDropdownCheckbox,
                items: DataSource.SensitivityLabelItems.slice(1),
                label: "Find Files with Label:",
                placeholder: "Select Label(s)",
                description: "Filter results for specific sensitivity label(s)."
            } as Components.IFormControlPropsMultiDropdownCheckbox
        ];
    }

    // Analyzes the libraries
    private static analyzeLibraries(webId: string, webUrl: string, libraries: Types.SP.ListOData[], drives: Types.Microsoft.Graph.drive[], withLabelsFl, withoutLabelsFl) {
        // Return a promise
        return new Promise(resolve => {
            let counter = 0;
            let siteText = this._elSubNav.children[0].innerHTML;

            // Parse the libraries
            Helper.Executor(libraries, lib => {
                let fileItems: ISensitivityLabelItem[] = [];

                // See if we are stopping this process
                if (this._stopFl) { return; }

                // Update the dialog
                this._elSubNav.children[0].innerHTML = `${siteText} [Analyzing Library ${++counter} of ${libraries.length}]: ${lib.Title}`;

                // Get the drive for this library
                let drive = drives.find(drive => {
                    return drive.name == lib.Title || drive.webUrl.endsWith(lib.RootFolder.ServerRelativeUrl);
                });

                // Ensure a drive exists, otherwise check the next library
                if (drive == null) { return; }

                // Return a promise
                return new Promise(resolve => {
                    // Update the dialog
                    this._elSubNav.children[1].innerHTML = `Analyzing the files for this library...`;

                    // Get the files for this library
                    let filesProcessed = 0;
                    DataSource.loadFiles(webId, webUrl, drive.id, null, true, (file: Types.Microsoft.Graph.driveItem) => {
                        let hasLabel = file.sensitivityLabel && file.sensitivityLabel.displayName ? true : false;

                        // Update the dialog
                        this._elSubNav.children[1].innerHTML = `Analyzing the files for this library. Files Analyzed: ${++filesProcessed}`;

                        // Add the file, based on the flags
                        if ((withLabelsFl && hasLabel) || (withoutLabelsFl && !hasLabel)) {
                            // See if we are filter for a label
                            if (this._filterLabels.length > 0) {
                                // See if this is a target label
                                if (this._filterLabels.indexOf(file.sensitivityLabel.id) < 0) { return; }
                            }

                            let fileInfo = file.name.split('.');
                            let folderPath = file.parentReference.path.split('root:')[1];

                            // Append the data
                            let fileItem: ISensitivityLabelItem = {
                                Author: file.createdBy.user["email"] || file.createdBy.user["displayName"],
                                File: file,
                                FileExtension: fileInfo[fileInfo.length - 1],
                                FileName: file.name,
                                HasUniquePermissions: file.listItem["HasUniquePermissions"],
                                ItemId: file.listItem["Id"],
                                ListId: lib.Id,
                                ListTitle: lib.Title,
                                Overshared: ViewPermissions.isOvershared(file.listItem["RoleAssignments"].results) ? "Yes" : "No",
                                Path: `${lib.RootFolder.ServerRelativeUrl}${folderPath}/${file.name}`,
                                Permissions: file.listItem["RoleAssignments"].results,
                                SensitivityLabel: file.sensitivityLabel.displayName,
                                SensitivityLabelId: file.sensitivityLabel.id,
                                WebId: webId,
                                WebUrl: webUrl
                            };

                            // Save a reference to the item
                            this._items.push(fileItem);
                            fileItems.push(fileItem);

                            // See if we have hit 100 items
                            if (fileItems.length >= 100) {
                                // Add the items to the datatable
                                this._dashboard.Datatable.addRow(fileItems);
                                fileItems = [];
                            }
                        }

                        // Return the stop flag
                        return this._stopFl;
                    }).then(() => {
                        // See if items exist
                        if (fileItems.length > 0) {
                            // Add the items to the datatable
                            this._dashboard.Datatable.addRow(fileItems);
                        }

                        // Check the next library
                        resolve(null);
                    });
                });
            }).then(resolve);
        });
    }

    // Refreshes the item
    private static refreshItem(dlpItem: ISensitivityLabelItem): PromiseLike<ISensitivityLabelItem> {
        // Return a promise
        return new Promise(resolve => {
            let web = this._loadOneDrive ? Web.getOneDrive() : Web(dlpItem.WebUrl, { requestDigest: DataSource.SiteContext.FormDigestValue });

            // Get the list item
            web.Lists(dlpItem.ListTitle).Items(dlpItem.ItemId).query({
                Expand: ["RoleAssignments/Member/Users", "RoleAssignments/RoleDefinitionBindings"],
                Select: ["Id", "HasUniqueRoleAssignments"]
            }).execute((item) => {
                // Update the item
                dlpItem.HasUniquePermissions = item.HasUniqueRoleAssignments;
                dlpItem.Overshared = ViewPermissions.isOvershared(item.RoleAssignments.results as any) ? "Yes" : "No";
                dlpItem.Permissions = item.RoleAssignments.results as any;

                // Find the item
                for (let i = 0; i < this._items.length; i++) {
                    let item = this._items[i];

                    // See if this is the item
                    if (item.ItemId == dlpItem.ItemId && item.ListId == dlpItem.ListId) {
                        // Update the item
                        this._items[i] = dlpItem;
                    }
                }

                // Resolve the request
                resolve(dlpItem);
            });
        });
    }

    // Renders the search summary
    private static renderSummary(el: HTMLElement, auditOnly: boolean, onClose: () => void) {
        // Render the summary
        this._dashboard = new Dashboard({
            el,
            navigation: {
                title: "Sensitivity Labels",
                showFilter: false,
                items: [{
                    text: "New Search",
                    className: "btn-outline-light",
                    isButton: true,
                    onClick: () => {
                        // Set the stop flag
                        this._stopFl = true;

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
                        new ExportCSV("SensitivityLabels.csv", CSVFields, this._items);
                    }
                }]
            },
            table: {
                rows: this._items,
                onRendering: dtProps => {
                    dtProps.columnDefs = [
                        {
                            "targets": [5, 6],
                            "orderable": false,
                            "searchable": false
                        }
                    ];

                    // Order by sensitivity label
                    dtProps.order = [[3, "asc"]];

                    // Return the properties
                    return dtProps;
                },
                columns: [
                    {
                        name: "ListTitle",
                        title: "List"
                    },
                    {
                        name: "FileName",
                        title: "File"
                    },
                    {
                        name: "Path",
                        title: "File Info",
                        onRenderCell: (el, col, item: ISensitivityLabelItem) => {
                            // Show the file info
                            el.innerHTML = `
                                <b>Created By: </b>${item.Author}
                                <br/>
                                <b>Path: </b>${item.Path}
                            `;
                        }
                    },
                    {
                        name: "SensitivityLabel",
                        title: "Sensitivity Label",
                        onRenderCell: (el) => {
                            // Make the text display in the middle
                            el.style.verticalAlign = "middle";
                        }
                    },
                    {
                        name: "",
                        title: "Overshared",
                        onRenderCell: (el, col, item: ISensitivityLabelItem) => {
                            let isOvershared = item.Overshared === "Yes" ? true : false;

                            // Set the order info
                            el.setAttribute("data-order", item.Overshared);

                            // Make the badge display in the middle
                            el.style.verticalAlign = "middle";

                            // Render a badge
                            let badge = Components.Badge({
                                el,
                                className: "me-2",
                                content: isOvershared ? "Overshared" : item.Overshared,
                                type: isOvershared ? Components.BadgeTypes.Danger : Components.BadgeTypes.Secondary,
                                isPill: true
                            });

                            // See if this is overshared
                            if (isOvershared) {
                                // Render a tooltip
                                Components.Tooltip({
                                    target: badge.el,
                                    content: `The file has been flagged as overshared because it's shared with the following groups:<br/>${ViewPermissions.getOversharedGroups(item.Permissions).join("<br/>")}`
                                });
                            }
                        }
                    },
                    {
                        name: "",
                        title: "Permissions",
                        onRenderCell: (el, col, item: ISensitivityLabelItem) => {
                            let adGroups = 0;
                            let m365Groups = 0;
                            let siteGroups = 0;
                            let users = 0;

                            // Parse the permissions
                            item.Permissions.forEach(role => {
                                // See if this is a user
                                switch (role.Member.PrincipalType) {
                                    case SPTypes.PrincipalTypes.User:
                                        users++;
                                        break;
                                    case SPTypes.PrincipalTypes.SharePointGroup:
                                        siteGroups++;
                                        break;
                                    default:
                                        let groupId = M365Groups.getGroupId(role.Member.LoginName);
                                        groupId ? m365Groups++ : adGroups++;
                                        break;
                                }
                            });

                            // Output the permission information
                            el.innerHTML = `
                                <b>Unique Permissions: </b>${item.HasUniquePermissions ? "Yes" : "No"}
                                <br/>
                                <b># of Users: </b>${users}
                                <br/>
                                <b># of Site Groups: </b>${siteGroups}
                                <br/>
                                <b># of AD Groups: </b>${adGroups}
                                <br/>
                                <b># of M365 Groups: </b>${m365Groups}
                                <br/>
                            `;
                        }
                    },
                    {
                        className: "text-end",
                        name: "",
                        title: "",
                        onRenderCell: (el, col, row: ISensitivityLabelItem, rowIdx) => {
                            // Render the buttons
                            let tooltips = Components.TooltipGroup({
                                el,
                                isSmall: true,
                                isVertical: true,
                                tooltips: [
                                    {
                                        content: "Click to view the document.",
                                        btnProps: {
                                            className: "pe-2 py-1",
                                            text: "View File",
                                            type: Components.ButtonTypes.OutlinePrimary,
                                            onClick: () => {
                                                // View the document
                                                window.open(Documents.isWopi(`${row.FileName}`) ? row.WebUrl + "/_layouts/15/WopiFrame.aspx?sourcedoc=" + row.Path + "&action=view" : row.Path, "_blank");
                                            }
                                        }
                                    },
                                    {
                                        content: "Click to view the permissions for this document.",
                                        btnProps: {
                                            className: "pe-2 py-1",
                                            text: "View Permissions",
                                            type: Components.ButtonTypes.OutlinePrimary,
                                            onClick: () => {
                                                // View the permissions for the document
                                                ViewPermissions.show(row);
                                            }
                                        }
                                    }
                                ]
                            });

                            // See if the file is overshared
                            if (!auditOnly && row.Overshared === "Yes") {
                                tooltips.add({
                                    content: "Click to remove the groups that are flagging this file as overshared.",
                                    btnProps: {
                                        className: "pe-2 py-1",
                                        text: "Secure File",
                                        type: Components.ButtonTypes.OutlinePrimary,
                                        onClick: () => {
                                            // Remove the overshared groups from the permissions
                                            ViewPermissions.removeOversharedGroups(row, () => {
                                                // Refresh the item
                                                this.refreshItem(row).then(updatedItem => {
                                                    // Update this row
                                                    this._dashboard.updateRow(rowIdx, updatedItem);
                                                });
                                            });
                                        }
                                    }
                                });
                            }

                            // Ensure we can make updates
                            if (!auditOnly) {
                                // Add the label button
                                tooltips.add({
                                    content: "Sets the label for the file.",
                                    btnProps: {
                                        className: "pe-2 py-1",
                                        text: "Set Label",
                                        type: Components.ButtonTypes.OutlinePrimary,
                                        onClick: () => {
                                            // Show the form to label the file
                                            this.showLabelFileForm(row.File, label => {
                                                // Update the row cell
                                                row.SensitivityLabel = label;
                                                this._dashboard.updateCell(rowIdx, 3, row);
                                            });
                                        }
                                    }
                                });
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
    }

    // Runs the report
    static run(el: HTMLElement, auditOnly: boolean, values: { [key: string]: string }, onClose: () => void) {
        let data: IWebItem[] = [];
        this._loadOneDrive = values["LoadOneDrive"] == "true";
        this._stopFl = false;

        // Clear the items
        this._items = [];

        // Set the flags
        let withLabelsFl = false;
        let withoutLabelsFl = false;
        (values["SearchType"] as any as Components.ICheckboxGroupItem[]).forEach(item => {
            withLabelsFl = item.name == "WithLabels" ? true : withLabelsFl;
            withoutLabelsFl = item.name == "WithoutLabels" ? true : withoutLabelsFl;
        });

        // Set the labels filter
        this._filterLabels = [];
        let labels = values["FilterLabel"] as any as Components.IDropdownItem[];
        (labels || []).forEach(label => {
            this._filterLabels.push(label.value);
        });

        // Show a loading dialog
        LoadingDialog.setHeader("Searching Site");
        LoadingDialog.setBody("Loading the libraries...");
        LoadingDialog.show();

        // Clear the element
        while (el.firstChild) { el.removeChild(el.firstChild); }

        // Render the summary
        this.renderSummary(el, auditOnly, onClose);

        // Hide the loading dialog
        LoadingDialog.hide();

        // Determine the webs to target
        let siteItems: Components.IDropdownItem[] = null;
        if (this._loadOneDrive) {
            siteItems = [{ text: DataSource.OneDriveWeb.Url, value: DataSource.OneDriveWeb.Id }] as any;
        } else {
            siteItems = values["TargetWeb"] && values["TargetWeb"]["value"] ? [values["TargetWeb"]] as any : DataSource.SiteItems;
        }

        // Parse the webs
        let counter = 0;
        Helper.Executor(siteItems, siteItem => {
            // See if we are stopping this process
            if (this._stopFl) { return; }

            // Update the status
            this._elSubNav.children[0].innerHTML = `Searching Site ${++counter} of ${siteItems.length}`;
            this._elSubNav.children[1].innerHTML = "Getting the libraries for this web...";

            // Return a promise
            return new Promise(resolve => {
                // Set the default filter
                let filter = `Hidden eq false and BaseTemplate eq ${SPTypes.ListTemplateType.DocumentLibrary} or BaseTemplate eq ${SPTypes.ListTemplateType.MySiteDocumentLibrary} or BaseTemplate eq ${SPTypes.ListTemplateType.PageLibrary}`;

                // See if a list was specified
                if (values["TargetList"]) {
                    // Set the filter
                    filter = `Title eq '${values["TargetList"]}'`;
                }

                // Get the libraries for this site
                let web = this._loadOneDrive ? Web.getOneDrive() : Web(siteItem.text, { requestDigest: DataSource.SiteContext.FormDigestValue });
                web.Lists().query({
                    Filter: filter,
                    Expand: ["RootFolder"],
                    GetAllItems: true,
                    Select: ["Id", "Title", "RootFolder/ServerRelativeUrl"],
                    Top: 5000
                }).execute(libs => {
                    // Add the libraries to analyze
                    libs.results.forEach(lib => {
                        // Append the item
                        data.push({
                            ListId: lib.Id,
                            ListTitle: lib.Title,
                            WebId: siteItem.value,
                            WebUrl: siteItem.text
                        });
                    });

                    // Update the dialog
                    this._elSubNav.children[1].innerHTML = "Loading the files for the libraries...";

                    // Get the drives for this web
                    v2.drives({
                        siteId: this._loadOneDrive ? DataSource.OneDriveSite.Id : DataSource.Site.Id,
                        webId: this._loadOneDrive ? DataSource.OneDriveWeb.Id : DataSource.Web.Id
                    }).execute(drives => {
                        // Analyze the libraries
                        this.analyzeLibraries(siteItem.value, siteItem.text, libs.results, drives.results, withLabelsFl, withoutLabelsFl).then(resolve);
                    });
                });
            });
        }).then(() => {
            // Hide the sub-nav
            this._elSubNav.classList.add("d-none");
        });
    }

    // Shows the report form for a library
    static runReportForLibrary(auditOnly: boolean, values: { [key: string]: string }) {
        // Clear a modal form
        Modal.clear();
        Modal.setHeader("Sensitivity Files");
        Modal.setType(Components.ModalTypes.Full);

        // Run the report
        this.run(Modal.BodyElement, auditOnly, values, () => { });

        // Render the footer
        Components.ButtonGroup({
            el: Modal.FooterElement,
            className: "mt-3",
            buttons: [
                {
                    text: "Close",
                    type: Components.ButtonTypes.OutlineSecondary,
                    onClick: () => {
                        // Hide the form
                        Modal.hide();
                    }
                }
            ]
        });

        // Show the form
        Modal.show();
    }

    // Shows the form to label a file
    private static showLabelFileForm(file: Types.Microsoft.Graph.driveItem, onUpdate: (label: string) => void) {
        // Set the canvas
        CanvasForm.clear();
        CanvasForm.setHeader("Set Sensitivity Label");
        CanvasForm.setSize(Components.OffcanvasSize.Medium2);
        CanvasForm.setType(Components.OffcanvasTypes.End);

        // Set the content
        CanvasForm.setBody(`
            <div></div>
            <div class="d-flex justify-content-end"></div>
        `);

        // Set the form
        let form = Components.Form({
            el: CanvasForm.BodyElement.querySelector("div"),
            groupClassName: "mb-3",
            controls: [
                {
                    name: "SensitivityLabel",
                    label: "Select Sensitivity Label:",
                    description: "This will set any file that isn't currently labelled.",
                    errorMessage: "A sensitivity label is required.",
                    items: DataSource.SensitivityLabelItems,
                    type: Components.FormControlTypes.Dropdown,
                    required: true,
                    onValidate: (ctrl, results) => {
                        // Ensure a selection exists
                        results.isValid = results.value && results.value.text ? true : false;
                        return results;
                    }
                } as Components.IFormControlPropsDropdown,
                {
                    name: "Justification",
                    label: "Justification:",
                    description: "Your organization requires justification to change this label.",
                    type: Components.FormControlTypes.Dropdown,
                    required: true,
                    items: [
                        { text: "Previous label no longer applies" },
                        { text: "Previous label was incorrect" },
                        { text: "Other" }
                    ],
                    onChange: (item) => {
                        let ctrlTextbox = form.getControl("JustificationOther");

                        // See if we are showing it
                        if (item.text == "Other") {
                            // Show it
                            ctrlTextbox.show();
                        } else {
                            // Hide it
                            ctrlTextbox.hide();
                        }
                    }
                } as Components.IFormControlPropsDropdown,
                {
                    name: "JustificationOther",
                    label: "Explain Justification:",
                    description: "Do not enter sensitive information",
                    className: "d-none",
                    type: Components.FormControlTypes.TextField,
                    errorMessage: "A justification is required.",
                    onValidate: (ctrl, results) => {
                        let item = form.getValues()["Justification"] as Components.IDropdownItem;

                        // See if we are expecting a justification
                        if (item.text == "Other") {
                            // Set the falg
                            results.isValid = results.value ? true : false;
                        }

                        // Return the results
                        return results;
                    }
                } as Components.IFormControlPropsTextField,
            ]
        });

        // Set the footer
        Components.TooltipGroup({
            el: CanvasForm.BodyElement.querySelector("div.d-flex"),
            tooltips: [
                {
                    content: "Sets the default sensitivity label to the selected option.",
                    btnProps: {
                        text: "Update",
                        type: Components.ButtonTypes.OutlinePrimary,
                        onClick: () => {
                            // Ensure the form is valid
                            if (form.isValid()) {
                                let values = form.getValues();
                                let label: Components.IDropdownItem = values["SensitivityLabel"];

                                // Update the justification
                                let justification = values["Justification"].text;
                                justification = justification == "Other" ? values["JustificationOther"] : justification;

                                // Show a loading dialog
                                LoadingDialog.setHeader("Setting Label");
                                LoadingDialog.setBody("Updating the label for this file.");
                                LoadingDialog.show();

                                // Label the file
                                BulkLabel.labelFile(file, label.text, label.value, justification, []).then(response => {
                                    // See if it was successful
                                    if (!response.errorFl) {
                                        // Call the event
                                        onUpdate(label.text);

                                        // Hide the dialog
                                        CanvasForm.hide();
                                    } else {
                                        // Set the error
                                        let ctrl = form.getControl("SensitivityLabel");
                                        ctrl.updateValidation(ctrl.el, {
                                            isValid: false,
                                            invalidMessage: response.message
                                        });
                                    }

                                    // Hide the dialog
                                    LoadingDialog.hide();
                                });
                            }
                        }
                    }
                },
                {
                    content: "Closes the dialog.",
                    btnProps: {
                        text: "Close",
                        type: Components.ButtonTypes.OutlineSecondary,
                        onClick: () => {
                            // Close the form
                            CanvasForm.hide();
                        }
                    }
                }
            ]
        });

        // Show the form
        CanvasForm.show();
    }

    // Stops the report
    static stop() {
        // Set the stop flags
        this._stopFl = true;
        BulkLabel.stop();
    }
}