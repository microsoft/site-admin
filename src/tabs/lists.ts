import { Dashboard, Modal, LoadingDialog } from "dattatable";
import { Components, Helper, SPTypes, Types, v2, Web } from "gd-sprest-bs";
import { ExportCSV } from "../reports/exportCSV";
import { IAppProps } from "../app";
import { DataSource } from "../ds";
import { ReportTypes } from "./reports";
import { DLP } from "../reports/dlp";
import { SearchEEEU } from "../reports/searchEEEU";
import { SensitivityLabels } from "../reports/sensitivityLabels";

interface IList {
    BaseTemplate: number;
    DefaultSensitivityLabel: string;
    DriveId: string;
    HasUniqueRoleAssignments: boolean;
    IncludedInSearch: boolean;
    ItemCount: number;
    ListId: string;
    ListName: string;
    ListTemplateType: number;
    ListTemplate: string;
    ListUrl: string;
    ListViewUrl: string;
    WebId: string;
    WebUrl: string;
}

const CSVFields = [
    "WebUrl", "WebId", "ListName", "ListTemplateType", "ListTemplate", "ListUrl", "ListViewUrl", "HasUniqueRoleAssignments", "ItemCount", "IncludedInSearch", "DefaultSensitivityLabel"
]

/**
 * Lists Tab
 */
export class ListsTab {
    private _appProps: IAppProps = null;
    private _dt: Dashboard = null;
    private _el: HTMLElement = null;
    private _elSubNav: HTMLElement = null;
    private _items: IList[] = [];

    // List Template Types
    private static _listTemplates: { [key: number]: string } = {};

    // Constructor
    constructor(el: HTMLElement, appProps: IAppProps) {
        this._el = el;
        this._appProps = appProps;

        // Render the tab
        this.render();
    }

    // Analyzes a list
    private analyzeList(webUrl: string, webId: string, list: Types.SP.ListOData, drives: Types.Microsoft.Graph.drive[]) {
        // See if a drive exists for this list
        let drive = drives.find(drive => {
            return drive.name == list.Title || drive.webUrl.endsWith(list.RootFolder.ServerRelativeUrl);
        });

        // Add a row for this entry
        let item: IList = {
            BaseTemplate: list.BaseTemplate,
            DefaultSensitivityLabel: list.DefaultSensitivityLabelForLibrary,
            DriveId: drive?.id,
            HasUniqueRoleAssignments: list.HasUniqueRoleAssignments,
            IncludedInSearch: !list.NoCrawl,
            ItemCount: list.ItemCount,
            ListId: list.Id,
            ListName: list.Title,
            ListTemplate: ListsTab._listTemplates[list.BaseTemplate],
            ListTemplateType: list.BaseTemplate,
            ListUrl: list.RootFolder.ServerRelativeUrl,
            ListViewUrl: list.DefaultViewUrl,
            WebId: webId,
            WebUrl: webUrl
        };
        this._items.push(item);
        this._dt.Datatable.addRow(item);
    }

    // Loads the lists
    loadLists(showHiddenLists: boolean = false): PromiseLike<void> {
        // Update the sub-nav
        this._elSubNav.classList.remove("d-none");
        this._elSubNav.children[0].innerHTML = "Loading lists for web: " + DataSource.Web.Url;

        // Clear the items
        this._items = [];

        // Clear the table
        this._dt.refresh(this._items);

        // Return a promise
        return new Promise(resolve => {
            // Set the filter
            let Filter = showHiddenLists ? "" : "Hidden eq false";

            // Get the list templates
            let web = Web(DataSource.Web.Url, { requestDigest: DataSource.SiteContext.FormDigestValue });
            web.ListTemplates().execute(templates => {
                // Parse the templates
                for (let i = 0; i < templates.results.length; i++) {
                    let template = templates.results[i];

                    // Skip the template if we already set it
                    if (ListsTab._listTemplates[template.ListTemplateTypeKind]) { continue; }

                    // Set the template
                    ListsTab._listTemplates[template.ListTemplateTypeKind] = template.Name;
                }
            });

            // Get the lists
            web.Lists().query({
                Filter,
                Expand: ["DefaultViewFormUrl", "RootFolder"],
                OrderBy: ["Title"],
                Select: [
                    "BaseTemplate", "DefaultSensitivityLabelForLibrary", "Id", "NoCrawl",
                    "ItemCount", "Title", "HasUniqueRoleAssignments", "RootFolder/ServerRelativeUrl"
                ]
            }).execute(lists => {
                let ctrList = 0;

                // Get the drives for this web
                v2.drives({ siteId: DataSource.Site.Id, webId: DataSource.Web.Id }).execute(drives => {
                    // Parse the lists
                    Helper.Executor(lists.results, list => {
                        // Update the status
                        this._elSubNav.children[0].innerHTML = `Analyzing List ${++ctrList} of ${lists.results.length}...`;

                        // Analyze the list
                        return this.analyzeList(DataSource.Web.Url, DataSource.Web.Id, list, drives.results);
                    }).then(() => {
                        // Hide the sub-nav
                        this._elSubNav.classList.add("d-none");

                        // Resolve the request
                        resolve();
                    });
                });
            }, true);
        });
    }

    // Generate the actions
    private generateActions(el: HTMLElement, item: IList, rowIdx: number) {
        let isLibrary = item.ListTemplateType == SPTypes.ListTemplateType.DocumentLibrary ||
            item.ListTemplateType == SPTypes.ListTemplateType.MySiteDocumentLibrary ||
            item.ListTemplateType == SPTypes.ListTemplateType.WebPageLibrary

        // Render the buttons
        let tooltips = Components.TooltipGroup({
            el,
            tooltips: [
                {
                    content: "Click to view the list.",
                    btnProps: {
                        text: "View List",
                        type: Components.ButtonTypes.OutlinePrimary,
                        onClick: () => {
                            // Show the list/library
                            window.open(item.ListUrl, "_blank");
                        }
                    }
                }
            ]
        });

        // Add the options to make changes
        if (!this._appProps.auditOnly || !this._appProps.hideReports?.sensitivityLabels) {
            let tooltip: Components.ITooltip = null;

            // Ensure this is a library and has a drive
            if (isLibrary && item.DriveId) {
                // Set the default label
                tooltips.add({
                    content: "Click to set the default sensitivity label.",
                    btnProps: {
                        isDisabled: !DataSource.HasSensitivityLabels,
                        text: "Default Label",
                        type: Components.ButtonTypes.OutlinePrimary,
                        onClick: () => {
                            // Show the form
                            this.setDefaultSensitivityLabel(item);
                        }
                    }
                });
            }

            // Update the list to be in or out of the search index
            tooltips.add({
                assignTo: obj => { tooltip = obj; },
                content: `Click to ${item.IncludedInSearch ? "remove" : "add"} the content from the search index.`,
                btnProps: {
                    text: "Update Search",
                    type: Components.ButtonTypes.OutlinePrimary,
                    onClick: () => {
                        // Show the form
                        this.setListSearch(item, () => {
                            // Flip the flag
                            item.IncludedInSearch = !item.IncludedInSearch;
                            tooltip.setContent(`Click to ${item.IncludedInSearch ? "remove" : "add"} the content from the search index.`);

                            // Update the row cell
                            this._dt.updateCell(rowIdx, 4, item.IncludedInSearch);
                        });
                    }
                }
            });
        }

        // Add the reports tooltip
        tooltips.add({
            content: "Click to run an audit report on this list.",
            btnProps: {
                text: "View Reports",
                type: Components.ButtonTypes.OutlinePrimary,
                onClick: () => {
                    // Show the reports form
                    this.showReportsForm(isLibrary, item);
                }
            }
        });
    }

    // Renders the tab
    private render() {
        // Clear the element
        while (this._el.firstChild) { this._el.removeChild(this._el.firstChild); }

        // Render a dashboard
        this._dt = new Dashboard({
            el: this._el,
            navigation: {
                title: "Lists/Libraries",
                showFilter: false,
                itemsEnd: [{
                    text: "Export to CSV",
                    className: "btn-outline-light me-2",
                    isButton: true,
                    onClick: () => {
                        // Export the CSV
                        new ExportCSV("listPermissions.csv", CSVFields, this._items);
                    }
                }]
            },
            table: {
                rows: this._items,
                onRendering: dtProps => {
                    dtProps.columnDefs = [
                        {
                            "targets": 5,
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
                        name: "ListName",
                        title: "List Info",
                        onRenderCell: (el, col, item: IList) => {
                            // Render the info
                            el.innerHTML = `
                                <b>Name: </b>${item.ListName}
                                <br/>
                                <b>Type: </b>${item.ListTemplate}
                                <br/>
                                <b>Url: </b>${item.ListUrl}
                            `;
                        }
                    },
                    {
                        name: "ItemCount",
                        title: "Item Count"
                    },
                    {
                        name: "DefaultSensitivityLabel",
                        title: "Default Sensitivity Label",
                        onRenderCell: (el, col, item: IList) => {
                            // See if a value exists
                            if (item.DefaultSensitivityLabel) {
                                // Set the value
                                el.innerHTML = DataSource.getSensitivityLabel(item.DefaultSensitivityLabel);
                            }
                        }
                    },
                    {
                        name: "IncludedInSearch",
                        title: "Included In Search"
                    },
                    {
                        name: "HasUniqueRoleAssignments",
                        title: "Has Unique Permissions"
                    },
                    {
                        className: "text-end",
                        name: "",
                        title: "",
                        onRenderCell: (el, col, item: IList, rowIdx) => {
                            // Generate the actions
                            this.generateActions(el, item, rowIdx);
                        }
                    }
                ]
            }
        });

        // Set the sub-nav element
        this._elSubNav = this._el.querySelector("#sub-navigation");
        this._elSubNav.classList.remove("d-none");
        this._elSubNav.classList.add("my-2");
        this._elSubNav.innerHTML = `<div class="h6">Loading the webs...</div>`;
    }

    // Reverts the item permissions
    private setDefaultSensitivityLabel(item: IList) {
        // Set the modal header
        Modal.clear();
        Modal.setHeader("Set Default Sensitivity Label");

        // Set the form
        let form = Components.Form({
            el: Modal.BodyElement,
            controls: [
                {
                    name: "SensitivityLabel",
                    label: "Select Sensitivity Label:",
                    description: "This will set the default sensitivity label for this library.",
                    items: DataSource.SensitivityLabelItems,
                    type: Components.FormControlTypes.Dropdown,
                    required: true,
                    value: item.DefaultSensitivityLabel
                } as Components.IFormControlPropsDropdown
            ]
        });

        // Set the footmer
        Components.TooltipGroup({
            el: Modal.FooterElement,
            tooltips: [
                {
                    content: "Sets the default sensitivity label to the selected option.",
                    btnProps: {
                        text: "Update",
                        type: Components.ButtonTypes.OutlinePrimary,
                        onClick: () => {
                            // Ensure the form is valid
                            if (form.isValid()) {
                                let labelId = form.getValues()["SensitivityLabel"].value;

                                // Show a loading dialog
                                LoadingDialog.setHeader("Updating List");
                                LoadingDialog.setBody("This dialog will close after the list is updated...");
                                LoadingDialog.show();

                                // Set the logic to run after the update completes
                                let onComplete = () => {
                                    // Hide the dialogs
                                    LoadingDialog.hide();
                                    Modal.hide();
                                }

                                // Restore the permissions
                                Web(item.WebUrl, { requestDigest: DataSource.SiteContext.FormDigestValue })
                                    .Lists().getById(item.ListId).update({
                                        DefaultSensitivityLabelForLibrary: labelId
                                    }).execute(() => {
                                        // Update the item
                                        item.DefaultSensitivityLabel = labelId;

                                        // Update the data table
                                        // TODO

                                        // Run the complete logic
                                        onComplete();
                                    }, onComplete);
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
                            // Close the modal
                            Modal.hide();
                        }
                    }
                }
            ]
        });

        // Show the modal
        Modal.show();
    }

    // Adds/Removes the list content from the search index
    private setListSearch(item: IList, onUpdate: () => void) {
        // Clear the modal
        Modal.clear();

        // Set the header
        Modal.setHeader(item.IncludedInSearch ? "Remove From Search" : "Add To Search");

        // Set the body
        Modal.setBody(item.IncludedInSearch ? "Confirm to remove the contents of this list from the search index." : "Confirm to add the contents of this list to the search index.");

        // Set the footer
        Components.TooltipGroup({
            el: Modal.FooterElement,
            tooltips: [
                {
                    content: item.IncludedInSearch ? "Remove the list from search" : "Add the list to search",
                    btnProps: {
                        text: "Update",
                        type: Components.ButtonTypes.OutlinePrimary,
                        onClick: () => {
                            // Show a loading dialog
                            LoadingDialog.setHeader("Update List");
                            LoadingDialog.setBody("This will close after the list has been updated...");
                            LoadingDialog.show();

                            // Update the list
                            Web(item.WebUrl, { requestDigest: DataSource.SiteContext.FormDigestValue }).Lists().getById(item.ListId).update({
                                NoCrawl: item.IncludedInSearch
                            }).execute(() => {
                                // Call the update event
                                onUpdate();

                                // Close the dialog
                                LoadingDialog.hide();
                                Modal.hide();
                            }, () => {
                                // Error updating the list
                                console.error("Error updating the search index for the list...");
                                LoadingDialog.hide();
                                Modal.hide();
                            });
                        }
                    }
                },
                {
                    content: "Closes the dialog.",
                    btnProps: {
                        text: "Close",
                        type: Components.ButtonTypes.OutlineSecondary,
                        onClick: () => {
                            // Close the modal
                            Modal.hide();
                        }
                    }
                }
            ]
        });

        // Show the modal
        Modal.show();
    }

    // Generates the tooltip options for the reports
    private showReportsForm(isLibrary: boolean, item: IList) {
        // Set the available reports
        let items: Components.IDropdownItem[] = [];

        // DLP
        if (typeof (this._appProps.hideReports.dlp) === "undefined" || this._appProps.hideReports.dlp != true) {
            items.push({
                text: "Data Loss Prevention",
                data: "Finds files that has DLP applied to it.",
                value: ReportTypes.DLP
            });
        }

        // Search EEEU
        if (typeof (this._appProps.hideReports.searchEEEU) === "undefined" || this._appProps.hideReports.searchEEEU != true) {
            items.push({
                text: "Search EEEU",
                data: "Search for the 'Every' and 'Everyone exception external users' accounts.",
                value: ReportTypes.SearchEEEU
            });
        }

        // Sensitivity Labels
        if (isLibrary && item.DriveId) {
            if (typeof (this._appProps.hideReports.sensitivityLabels) === "undefined" || this._appProps.hideReports.sensitivityLabels != true) {
                items.push({
                    text: "Sensitivity Labels",
                    data: "Click to view sensitivity label options.",
                    value: ReportTypes.SensitivityLabels,
                    onClick: () => {
                    }
                });
            }
        }

        // Unique Permissions
        if (typeof (this._appProps.hideReports.uniquePermissions) === "undefined" || this._appProps.hideReports.uniquePermissions != true) {
            items.push({
                text: "Unique Permissions",
                data: "Scans for items that have unique permissions.",
                value: ReportTypes.UniquePermissions
            });
        }

        // Clear the modal
        Modal.clear();
        Modal.setHeader("Select Report");

        // Set the body
        Components.Form({
            el: Modal.BodyElement,
            controls: [
                {
                    name: "ReportType",
                    label: "Select Report:",
                    description: "Select a report to run against this list.",
                    items,
                    type: Components.FormControlTypes.Dropdown,
                    required: true,
                    errorMessage: "A report selection is required.",
                    onChange: (selectedItem) => {
                        // Show the form for the selected report
                        switch (selectedItem?.value) {
                            case ReportTypes.DLP:
                                // Run the DLP report for this library
                                DLP.analyzeLibrary(item.WebId, item.WebUrl, item.ListId, item.ListName);
                                break;
                            case ReportTypes.SearchEEEU:
                                // Run the EEEU report for this library
                                SearchEEEU.searchList(item.WebUrl, item.ListName, this._appProps.auditOnly);
                                break;
                            case ReportTypes.SensitivityLabels:
                                // Load the folders for this list
                                DataSource.loadFolders(item.WebId, item.DriveId).then(folders => {
                                    // Show the senstivity label form
                                    SensitivityLabels.setDefaultSensitivityLabelForFiles(item.WebId, item.ListName, item.DriveId, item.DefaultSensitivityLabel, folders, this._appProps.disableSensitivityLabelOverride, this._appProps.reportProps.sensitivityLabelFileExt);
                                });
                                break;
                            case ReportTypes.UniquePermissions:
                                // TODO: Create a method to search this list for unique pemrissions
                                break;
                        }
                    }
                } as Components.IFormControlPropsDropdown
            ]
        });

        // Set the footer
        Components.TooltipGroup({
            el: Modal.FooterElement,
            tooltips: [
                {
                    content: "Closes the dialog.",
                    btnProps: {
                        text: "Close",
                        type: Components.ButtonTypes.OutlineSecondary,
                        onClick: () => {
                            // Close the modal
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