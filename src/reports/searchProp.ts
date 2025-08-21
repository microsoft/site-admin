import { Dashboard, LoadingDialog } from "dattatable";
import { Components, Search, Types } from "gd-sprest-bs";
import * as moment from "moment";
import { DataSource } from "../ds";
import Strings from "../strings";
import { ExportCSV } from "./exportCSV";

interface ISearchItem {
    LastModifiedTime: string;
    Path: string;
    Title: string;
    ViewsLast1Days: string;
    ViewsLast2Days: string;
    ViewsLast3Days: string;
    ViewsLast4Days: string;
    ViewsLast5Days: string;
    ViewsLast6Days: string;
    ViewsLast7Days: string;
    ViewsLastMonths1: string;
    ViewsLastMonths2: string;
    ViewsLastMonths3: string;
    ViewsLifeTime: string;
    ViewsRecent: string;
}

const CSVFields = [
    "Path", "Title", "LastModifiedTime", "ViewsLast1Days", "ViewsLast2Days", "ViewsLast3Days",
    "ViewsLast4Days", "ViewsLast5Days", "ViewsLast6Days", "ViewsLast7Days", "ViewsRecent",
    "ViewsLastMonths1", "ViewsLastMonths2", "ViewsLastMonths3", "ViewsLifeTime"
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
                    items: DataSource.SearchPropItems,
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
    private static renderSummary(el: HTMLElement, auditOnly: boolean, searchProp: string, items: ISearchItem[], onClose: () => void) {
        // Render the summary
        new Dashboard({
            el,
            navigation: {
                title: "Search Sites",
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
                    // Order by the 2nd column by default; ascending
                    dtProps.order = [[0, "asc"]];

                    // Return the properties
                    return dtProps;
                },
                columns: [
                    {
                        name: "",
                        title: "Site Information",
                        onRenderCell: (el, col, item: ISearchItem) => {
                            // Set the sort/filter values
                            el.setAttribute("data-filter", item.Path);
                            el.setAttribute("data-sort", item.Path);

                            // Render the activity
                            el.innerHTML = `
                                <div><b>Title: </b><a href="#" target="_blank">${item.Title}</a></div>
                                <div><b>Url: </b>${item.Path}</div>
                                <div><b>${searchProp}: </b>${item[searchProp] || ""}</div>
                            `;

                            // Set the click event
                            el.querySelector("a").addEventListener("click", () => {
                                // Show the link in a new window
                                window.open(item.Path, "_blank");
                            });
                        }
                    },
                    {
                        name: "",
                        title: "Views (7 Days)",
                        onRenderCell: (el, col, item: ISearchItem) => {
                            // Set the value
                            el.innerHTML = [
                                item.ViewsLast1Days, item.ViewsLast2Days, item.ViewsLast3Days,
                                item.ViewsLast4Days, item.ViewsLast5Days, item.ViewsLast6Days,
                                item.ViewsLast7Days].reduce((value, sum) => { return (parseInt(value) + parseInt(sum)).toString(); });
                        }
                    },
                    {
                        name: "",
                        title: "Views (14 Days)",
                        onRenderCell: (el, col, item: ISearchItem) => {
                            // Set the value
                            el.innerHTML = item.ViewsRecent || "0";
                        }
                    },
                    {
                        name: "ViewsLastMonths1",
                        title: "Views (30 Days)"
                    },
                    {
                        name: "",
                        title: "Views (60 Days)",
                        onRenderCell: (el, col, item: ISearchItem) => {
                            // Set the value
                            el.innerHTML = [
                                item.ViewsLastMonths1, item.ViewsLastMonths2
                            ].reduce((value, sum) => { return (parseInt(value) + parseInt(sum)).toString(); });
                        }
                    },
                    {
                        name: "",
                        title: "Views (90 Days)",
                        onRenderCell: (el, col, item: ISearchItem) => {
                            // Set the value
                            el.innerHTML = [
                                item.ViewsLastMonths1, item.ViewsLastMonths2, item.ViewsLastMonths3
                            ].reduce((value, sum) => { return (parseInt(value) + parseInt(sum)).toString(); });
                        }
                    },
                    {
                        name: "",
                        title: "Life Time Views",
                        onRenderCell: (el, col, item: ISearchItem) => {
                            // Set the value
                            el.innerHTML = item.ViewsLifeTime || "0";
                        }
                    },
                    {
                        name: "",
                        title: "Last Modified Time",
                        onRenderCell: (el, col, item: ISearchItem) => {
                            // Render the date
                            el.innerHTML = item.LastModifiedTime ? moment(item.LastModifiedTime).format(Strings.TimeFormat) : "";
                        }
                    }
                ]
            }
        });
    }

    // Runs the report
    static run(el: HTMLElement, auditOnly: boolean, searchProp: string, searchValue: string, onClose: () => void) {
        let searchEmpty = searchValue ? false : true;

        // Show a loading dialog
        LoadingDialog.setHeader("Searching Tenant");
        LoadingDialog.setBody(searchValue ? "Searching the sites tagged with '" + searchValue + "'..." : "Searching for sites not tagged...");
        LoadingDialog.show();

        // Set the query
        let query: Types.Microsoft.Office.Server.Search.REST.SearchRequest = {
            Querytext: `contentclass=sts_site AND ${searchProp}${searchEmpty ? "<>\"*\"" : ":\"" + searchValue + "\""}`,
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
            // Clear the element
            while (el.firstChild) { el.removeChild(el.firstChild); }

            // See if we are searching for empty values
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

            // Render the summary
            this.renderSummary(el, auditOnly, searchProp, search.results, onClose);

            // Hide the loading dialog
            LoadingDialog.hide();
        });
    }
}