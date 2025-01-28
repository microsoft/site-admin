import { Dashboard, Documents, LoadingDialog } from "dattatable";
import { Components, Search, Web } from "gd-sprest-bs";
import { fileEarmark } from "gd-sprest-bs/build/icons/svgs/fileEarmark";
import { fileEarmarkArrowDown } from "gd-sprest-bs/build/icons/svgs/fileEarmarkArrowDown";
import { trash } from "gd-sprest-bs/build/icons/svgs/trash";
import * as moment from "moment";
import { DataSource } from "../ds";
import Strings from "../strings";
import { ExportCSV } from "./exportCSV";

interface ISearchItem {
    Author: string;
    FileExtension: string;
    HitHighlightedSummary: string;
    LastModifiedTime: string;
    ListId: string;
    Path: string;
    SPSiteUrl: string;
    SPWebUrl: string;
    Title: string;
    WebId: string;
}

const CSVFields = [
    "Author", "FileExtension", "HitHighlightedSummary", "LastModifiedTime",
    "ListId", "Path", "SPSiteUrl", "SPWebUrl", "Title", "WebId"
]

export class SearchDocs {
    // Deletes a document
    private static deleteDocument(item: ISearchItem) {
        // Display a loading dialog
        LoadingDialog.setHeader("Deleting Document");
        LoadingDialog.setBody("Deleting the document '" + item.Title + "'. This will close after the request completes.");
        LoadingDialog.show();

        // Delete the document
        Web(item.SPWebUrl, { requestDigest: DataSource.SiteContext.FormDigestValue }).getFileByServerRelativeUrl(item.Path).delete().execute(
            // Success
            () => {
                // TODO - Display the confirmation

                // Close the dialog
                LoadingDialog.hide();
            },

            // Error
            () => {
                // TODO - Display an error

                // Close the dialog
                LoadingDialog.hide();
            }
        )
    }

    // Gets the form fields to display
    static getFormFields(): Components.IFormControlProps[] {
        return [
            {
                label: "Search Terms",
                name: "SearchTerms",
                className: "mb-3",
                description: "Enter the search terms using quotes for phrases [Ex: movie \"social media\" show]",
                type: Components.FormControlTypes.TextField,
                required: true,
                value: Strings.SearchTerms
            },
            {
                label: "File Types",
                name: "FileTypes",
                className: "mb-3",
                type: Components.FormControlTypes.TextField,
                required: true,
                value: Strings.SearchFileTypes
            }
        ];
    }

    // Renders the search summary
    private static renderSummary(el: HTMLElement, items: ISearchItem[], onClose: () => void) {
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
                        new ExportCSV("searchDocs.csv", CSVFields, items);
                    }
                }]
            },
            table: {
                rows: items,
                onRendering: dtProps => {
                    dtProps.columnDefs = [
                        {
                            "targets": 6,
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
                        name: "ListId",
                        title: "List Id"
                    },
                    {
                        name: "Path",
                        title: "Document Url"
                    },
                    {
                        name: "Title",
                        title: "File Name",
                        onRenderCell: (el, col, item: ISearchItem) => {
                            el.innerHTML = item.Title + "." + item.FileExtension;
                        }
                    },
                    {
                        name: "Author",
                        title: "Author(s)",
                        onRenderCell: (el, col, item: ISearchItem) => {
                            // Clear the cell
                            el.innerHTML = "";

                            // Validate Author exists & split by ;
                            let authors = (item.Author && item.Author.split(";")) || [item.Author];

                            // Parse the Authors
                            authors.forEach(author => {
                                // Append the Author
                                el.innerHTML += (author + "<br/>");
                            });
                        }
                    },
                    {
                        name: "LastModifiedTime",
                        title: "Modified",
                        onRenderCell: (el, col, item: ISearchItem) => {
                            el.innerHTML = item.LastModifiedTime ? moment(item.LastModifiedTime).format(Strings.TimeFormat) : "";
                        }
                    },
                    {
                        name: "",
                        title: "Search Result",
                        onRenderCell: (el, col, item: ISearchItem) => {
                            // Add the data-filter attribute for searching notes properly
                            el.setAttribute("data-filter", item.HitHighlightedSummary);

                            // Add the data-order attribute for sorting notes properly
                            el.setAttribute("data-order", item.HitHighlightedSummary);

                            // Declare a span element
                            let span = document.createElement("span");

                            // Return the plain text if less than 50 chars
                            if (el.innerHTML.length < 50) {
                                span.innerHTML = item.HitHighlightedSummary;
                            } else {
                                // Truncate to the last white space character in the text after 50 chars and add an ellipsis
                                span.innerHTML = item.HitHighlightedSummary.substring(0, 50).replace(/\s([^\s]*)$/, '') + '&#8230';

                                // Add a tooltip containing the text
                                Components.Tooltip({
                                    content: "<small>" + item.HitHighlightedSummary + "</small>",
                                    target: span
                                });
                            }

                            // Append the span
                            el.appendChild(span);
                        }
                    },
                    {
                        className: "text-end",
                        name: "",
                        title: "",
                        onRenderCell: (el, col, item: ISearchItem) => {
                            let btnDelete: Components.IButton = null;

                            // Render the buttons
                            Components.TooltipGroup({
                                el,
                                tooltips: [
                                    {
                                        content: "View Document",
                                        btnProps: {
                                            className: "pe-2 py-1",
                                            iconClassName: "mx-1",
                                            iconType: fileEarmark,
                                            iconSize: 24,
                                            text: "View",
                                            type: Components.ButtonTypes.OutlinePrimary,
                                            onClick: () => {
                                                // Show the security group
                                                window.open(Documents.isWopi(`${item.Title}.${item.FileExtension}`) ? item.SPWebUrl + "/_layouts/15/WopiFrame.aspx?sourcedoc=" + item.Path + "&action=view" : item.Path, "_blank");
                                            }
                                        }
                                    },
                                    {
                                        content: "Download Document",
                                        btnProps: {
                                            className: "pe-2 py-1",
                                            iconClassName: "mx-1",
                                            iconType: fileEarmarkArrowDown,
                                            iconSize: 24,
                                            text: "Download",
                                            type: Components.ButtonTypes.OutlinePrimary,
                                            onClick: () => {
                                                // Download the document
                                                window.open(`${item.SPWebUrl}/_layouts/15/download.aspx?SourceUrl=${item.Path}`, "_blank");
                                            }
                                        }
                                    },
                                    {
                                        content: "Delete Document",
                                        btnProps: {
                                            assignTo: btn => { btnDelete = btn; },
                                            className: "pe-2 py-1",
                                            iconClassName: "mx-1",
                                            iconType: trash,
                                            iconSize: 24,
                                            text: "Delete",
                                            type: Components.ButtonTypes.OutlineDanger,
                                            onClick: () => {
                                                // Confirm the deletion of the group
                                                if (confirm("Are you sure you want to delete this document?")) {
                                                    // Disable this button
                                                    btnDelete.disable();

                                                    // Delete the document
                                                    this.deleteDocument(item);
                                                }
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
        LoadingDialog.setHeader("Searching Site");
        LoadingDialog.setBody("Searching the content on this site...");
        LoadingDialog.show();

        // Get the form values
        let fileExt = (values["FileTypes"]).split(' ');
        let searchTerms = (values["SearchTerms"] || "").split(' ');

        Search.postQuery({
            url: DataSource.SiteContext.SiteFullUrl,
            targetInfo: { requestDigest: DataSource.SiteContext.FormDigestValue },
            query: {
                Querytext: `${searchTerms.join(" OR ")} IsDocument: true path: ${DataSource.SiteContext.SiteFullUrl}`,
                RefinementFilters: {
                    results: [`fileExtension:or("${fileExt.join('", "')}")`]
                },
                RowLimit: 500,
                SelectProperties: {
                    results: [
                        "Author", "FileExtension", "HitHighlightedSummary", "LastModifiedTime",
                        "ListId", "Path", "SPSiteUrl", "SPWebUrl", "Title", "WebId"
                    ]
                }
            }
        }).then(search => {
            // Clear the element
            while (el.firstChild) { el.removeChild(el.firstChild); }

            // Render the summary
            this.renderSummary(el, search.results, onClose);

            // Hide the loading dialog
            LoadingDialog.hide();
        });
    }
}