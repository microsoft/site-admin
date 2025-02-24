import { Dashboard, Modal, LoadingDialog } from "dattatable";
import { Components, Helper, Types, Web } from "gd-sprest-bs";
import { cardList } from "gd-sprest-bs/build/icons/svgs/cardList";
import { DataSource } from "../ds";
import { ExportCSV } from "./exportCSV";

interface IList {
    DefaultSensitivityLabel: string;
    HasUniqueRoleAssignments: boolean;
    ListName: string;
    ListTemplateType: number;
    ListTemplate: string;
    ListUrl: string;
    ListViewUrl: string;
    WebUrl: string;
}

const CSVFields = [
    "WebUrl", "ListType", "ListName", "ListUrl", "HasUniqueRoleAssignments", "DefaultSensitivityLabel"
]

export class Lists {
    private static _items: IList[] = [];

    // List Template Types
    private static _listTemplates: { [key: number]: string } = {};

    // Analyzes a list
    private static analyzeList(webUrl: string, list: Types.SP.ListOData) {
        // Add a row for this entry
        this._items.push({
            DefaultSensitivityLabel: list.DefaultSensitivityLabelForLibrary,
            HasUniqueRoleAssignments: list.HasUniqueRoleAssignments,
            ListName: list.Title,
            ListTemplate: this._listTemplates[list.BaseTemplate],
            ListTemplateType: list.BaseTemplate,
            ListUrl: list.RootFolder.ServerRelativeUrl,
            ListViewUrl: list.DefaultDisplayFormUrl,
            WebUrl: webUrl
        });
    }

    // Gets the form fields to display
    static getFormFields(): Components.IFormControlProps[] {
        return [
            {
                name: "ShowHiddenLists",
                title: "Show Hidden Lists?",
                type: Components.FormControlTypes.Switch,
                value: false
            }
        ];
    }

    // Renders the search summary
    private static renderSummary(el: HTMLElement, onClose: () => void) {
        // Render the summary
        new Dashboard({
            el,
            navigation: {
                title: "Search Content",
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
                        name: "WebUrl",
                        title: "Url"
                    },
                    {
                        name: "ListName",
                        title: "List Name"
                    },
                    {
                        name: "ListTemplate",
                        title: "List Template"
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
                        name: "HasUniqueRoleAssignments",
                        title: "Has Unique Permissions"
                    },
                    {
                        className: "text-end",
                        name: "",
                        title: "",
                        onRenderCell: (el, col, item: IList) => {
                            // Render the buttons
                            Components.TooltipGroup({
                                el,
                                tooltips: [
                                    {
                                        content: "Click to view the list.",
                                        btnProps: {
                                            className: "pe-2 py-1",
                                            iconClassName: "mx-1",
                                            iconSize: 24,
                                            iconType: cardList,
                                            text: "View",
                                            type: Components.ButtonTypes.OutlinePrimary,
                                            onClick: () => {
                                                // Show the security group
                                                window.open(item.ListViewUrl + "?ID=" + item.ListUrl, "_blank");
                                            }
                                        }
                                    },
                                    {
                                        content: "Click set the default sensitivity label.",
                                        btnProps: {
                                            className: "pe-2 py-1",
                                            //iconType: GetIcon(24, 24, "PeopleTeamDelete", "mx-1"),
                                            text: "Sensitivity Label",
                                            type: Components.ButtonTypes.OutlinePrimary,
                                            onClick: () => {
                                                // Show the form
                                                this.setDefaultSensitivityLabel(item);
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
    }

    // Runs the report
    static run(el: HTMLElement, values: { [key: string]: string }, onClose: () => void) {
        // Show a loading dialog
        LoadingDialog.setHeader("Searching Lists");
        LoadingDialog.setBody("Searching the site...");
        LoadingDialog.show();

        // Clear the items
        this._items = [];

        // See if we are showing hidden lists
        let showHiddenLists = values["ShowHiddenLists"];

        // Parse all the webs
        let counter = 0;
        Helper.Executor(DataSource.SiteItems, siteItem => {
            // Set the loading dialog element
            let elLoadingDialog = document.createElement("div");
            elLoadingDialog.classList.add("d-flex", "justify-content-center");
            elLoadingDialog.innerHTML = `<span>Search ${++counter} of ${DataSource.SiteItems.length}...</span><br/><span></span>`;
            let elStatus = elLoadingDialog.childNodes[2] as HTMLElement;

            // Update the loading dialog
            LoadingDialog.setBody(elLoadingDialog);

            // Return a promise
            return new Promise(resolve => {
                // Set the filter
                let Filter = showHiddenLists ? "" : "Hidden eq false";

                // Get the list templates
                let web = Web(siteItem.text, { requestDigest: DataSource.SiteContext.FormDigestValue });
                web.ListTemplates().execute(templates => {
                    // Parse the templates
                    for (let i = 0; i < templates.results.length; i++) {
                        let template = templates.results[i];

                        // Skip the template if we already set it
                        if (this._listTemplates[template.ListTemplateTypeKind]) { continue; }

                        // Set the template
                        this._listTemplates[template.ListTemplateTypeKind] = template.Name;
                    }
                });

                // Get the lists
                web.Lists().query({
                    Filter,
                    Expand: ["DefaultViewFormUrl", "RootFolder"],
                    Select: ["BaseTemplate", "DefaultSensitivityLabelForLibrary", "Id", "Title", "HasUniqueRoleAssignments", "RootFolder/ServerRelativeUrl"]
                }).execute(lists => {
                    let ctrList = 0;

                    // Parse the lists
                    Helper.Executor(lists.results, list => {
                        // Update the status
                        elStatus.innerHTML = `Analyzing List ${++ctrList} of ${lists.results.length}...`;

                        // Analyze the list
                        return this.analyzeList(siteItem.text, list);
                    }).then(resolve);
                }, true);
            });
        }).then(() => {
            // Clear the element
            while (el.firstChild) { el.removeChild(el.firstChild); }

            // Render the summary
            this.renderSummary(el, onClose);

            // Hide the loading dialog
            LoadingDialog.hide();
        });
    }

    // Reverts the item permissions
    private static setDefaultSensitivityLabel(item: IList) {
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
                                    .Lists(item.ListName).update({
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
}