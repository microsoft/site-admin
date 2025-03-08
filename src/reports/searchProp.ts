import { Dashboard, Documents, LoadingDialog } from "dattatable";
import { Components, Search, Types, Web } from "gd-sprest-bs";
import { fileEarmark } from "gd-sprest-bs/build/icons/svgs/fileEarmark";
import { fileEarmarkArrowDown } from "gd-sprest-bs/build/icons/svgs/fileEarmarkArrowDown";
import { trash } from "gd-sprest-bs/build/icons/svgs/trash";
import * as moment from "moment";
import { DataSource } from "../ds";
import { ExportCSV } from "./exportCSV";
import Strings from "../strings";

interface ISearchItem {
    Path: string;
    Title: string;
}

const CSVFields = [
    "Path", "Title"
]

export class SearchProp {
    // Gets the form fields to display
    static getFormFields(siteValue: string): Components.IFormControlProps[] {
        return [
            DataSource.SearchPropItems ?
                {
                    name: "value",
                    label: "Value:",
                    type: Components.FormControlTypes.Dropdown,
                    description: "Select a value to search for.",
                    items: [{ text: "", value: null } as Components.IDropdownItem].concat(DataSource.SearchPropItems),
                    value: siteValue
                } as Components.IFormControlPropsDropdown
                :
                {
                    name: "value",
                    label: "Value:",
                    type: Components.FormControlTypes.TextField,
                    description: "Enter a value to search for.",
                    value: siteValue
                }
        ];
    }

    // Renders the search summary
    private static renderSummary(el: HTMLElement, searchProp: string, items: ISearchItem[], onClose: () => void) {
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
                        new ExportCSV("searchProp.csv", CSVFields.concat([searchProp]), items);
                    }
                }]
            },
            table: {
                rows: items,
                onRendering: dtProps => {
                    dtProps.columnDefs = [
                        {
                            "targets": 3,
                            "orderable": false,
                            "searchable": false
                        }
                    ];

                    // Order by the 2nd column by default; ascending
                    dtProps.order = [[1, "asc"]];

                    // Return the properties
                    return dtProps;
                },
                columns: [
                    {
                        name: "Path",
                        title: "Site Url"
                    },
                    {
                        name: "Title",
                        title: "Site Title"
                    },
                    {
                        name: searchProp,
                        title: "Property Value"
                    },
                    {
                        className: "text-end",
                        name: "",
                        title: "",
                        onRenderCell: (el, col, item: ISearchItem) => {
                            // Render the buttons
                            Components.TooltipGroup({
                                el,
                                tooltips: [
                                    {
                                        content: "Click to view the site in a new tab.",
                                        btnProps: {
                                            text: "View Site",
                                            type: Components.ButtonTypes.OutlinePrimary,
                                            onClick: () => {
                                                // Show the site
                                                window.open(item.Path, "_blank");
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
    static run(el: HTMLElement, searchProp: string, searchValue: string, onClose: () => void) {
        let searchEmpty = searchValue ? false : true;

        // Show a loading dialog
        LoadingDialog.setHeader("Searching Tenant");
        LoadingDialog.setBody(searchValue ? "Searching the sites tagged with '" + searchValue + "'..." : "Searching for sites not tagged...");
        LoadingDialog.show();

        // Set the query
        let query: Types.Microsoft.Office.Server.Search.REST.SearchRequest = {
            Querytext: `contentclass=sts_site ${searchProp}${searchEmpty ? "<>" : "="}'${searchValue}'`,
            RowLimit: 500,
            SelectProperties: {
                results: CSVFields.concat([searchProp])
            }
        };

        // Search for the content
        Search.postQuery({
            getAllItems: true,
            url: DataSource.SiteContext.SiteFullUrl,
            targetInfo: { requestDigest: DataSource.SiteContext.FormDigestValue },
            query
        }).then(search => {
            // See if we are searching for a null value
            if (searchEmpty) {
                // Parse the results
                let results = [];
                for (let i = 0; i < search.results.length; i++) {
                    // Add the result if it doesn't exist
                    if (search.results[i][searchProp]) { continue; }
                    results.push(search.results[i]);
                }

                // Update the results
                search.results = results;
            }

            // Clear the element
            while (el.firstChild) { el.removeChild(el.firstChild); }

            // Render the summary
            this.renderSummary(el, searchProp, search.results, onClose);

            // Hide the loading dialog
            LoadingDialog.hide();
        });
    }
}