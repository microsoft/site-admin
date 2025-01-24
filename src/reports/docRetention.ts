import { DataTable, Documents, LoadingDialog } from "dattatable";
import { Components, Search, Web } from "gd-sprest-bs";
import { OfficeOnline } from "gd-sprest-bs/build/icons/custom/officeOnline";
import { fileEarmark } from "gd-sprest-bs/build/icons/svgs/fileEarmark";
import { fileEarmarkArrowDown } from "gd-sprest-bs/build/icons/svgs/fileEarmarkArrowDown";
import { trash } from "gd-sprest-bs/build/icons/svgs/trash";
import { xSquare } from "gd-sprest-bs/build/icons/svgs/xSquare";
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
    "Author",
    "FileExtension",
    "HitHighlightedSummary",
    "LastModifiedTime",
    "ListId",
    "Path",
    "SPSiteUrl",
    "SPWebUrl",
    "Title",
    "WebId"
]

export class DocRetention {
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

    // Renders the search summary
    private static renderSummary(el: HTMLElement, rows: ISearchItem[]) {
        // Render the table
        new DataTable({
            el,
            rows,
            onRendering: dtProps => {
                dtProps.columnDefs = [
                    {
                        "targets": 7,
                        "orderable": false,
                        "searchable": false
                    }
                ];

                // Order by the 4th column by default; ascending
                dtProps.order = [[3, "asc"]];

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
                    className: "text-end",
                    name: "",
                    title: "",
                    onRenderCell: (el, col, row: ISearchItem) => {
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
                                            // View the document
                                            window.open(Documents.isWopi(`${row.Title}.${row.FileExtension}`) ? row.SPWebUrl + "/_layouts/15/WopiFrame.aspx?sourcedoc=" + row.Path + "&action=view" : row.Path, "_blank");
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
                                            window.open(`${row.SPWebUrl}/_layouts/15/download.aspx?SourceUrl=${row.Path}`, "_blank");
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
                                                this.deleteDocument(row);
                                            }
                                        }
                                    }
                                }
                            ]
                        });
                    }
                }
            ]
        });
    }

    // Renders the footer
    private static renderFooter(el: HTMLElement, items: ISearchItem[], onClose: () => void) {
        // Set the footer
        let elFooter = document.createElement("div");
        elFooter.classList.add("d-flex align-items-end");
        el.appendChild(elFooter);

        // Render the buttons
        Components.TooltipGroup({
            el: elFooter,
            tooltips: [
                {
                    content: "Export to a CSV file",
                    btnProps: {
                        className: "pe-2 py-1",
                        iconType: OfficeOnline(24, 24, "mx-1"),
                        text: "Export",
                        type: Components.ButtonTypes.OutlineSuccess,
                        onClick: () => {
                            // Export the CSV
                            new ExportCSV("docRetention.csv", CSVFields, items);
                        }
                    }
                },
                {
                    content: "New Search",
                    btnProps: {
                        className: "pe-2 py-1",
                        iconClassName: "mx-1",
                        iconType: xSquare,
                        iconSize: 24,
                        text: "Close",
                        type: Components.ButtonTypes.OutlineSecondary,
                        onClick: () => {
                            // Call the close event
                            onClose();
                        }
                    }
                }
            ]
        });
    }

    // Runs the report
    static run(el: HTMLElement, startDate: string, onClose: () => void) {
        // Show a loading dialog
        LoadingDialog.setHeader("Searching Site");
        LoadingDialog.setBody("The site is being searched for files...");
        LoadingDialog.show();

        Search.postQuery<ISearchItem>({
            url: DataSource.SiteContext.SiteFullUrl,
            getAllItems: true,
            targetInfo: { requestDigest: DataSource.SiteContext.FormDigestValue },
            query: {
                Querytext: `IsDocument: true LastModifiedTime<${startDate} path: ${DataSource.SiteContext.SiteFullUrl}`,
                SelectProperties: {
                    results: [
                        "Author", "FileExtension", "HitHighlightedSummary", "LastModifiedTime",
                        "ListId", "Path", "SPSiteUrl", "SPWebUrl", "Title", "WebId"
                    ]
                }
            }
        }).then(search => {
            // Render the summary
            this.renderSummary(el, search.results);

            // Render the footer
            this.renderFooter(el, search.results, onClose);

            // Hide the loading dialog
            LoadingDialog.hide();
        });
    }
}