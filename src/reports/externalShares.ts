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
    "Author", "FileExtension", "ViewsLifeTime", "LastModifiedTime",
    "ListId", "Path", "SPSiteUrl", "SPWebUrl", "Title", "WebId"
]

export class ExternalShares {
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
    static getFormFields(): Components.IFormControlProps[] { return []; }

    // Renders the search summary
    private static renderSummary(el: HTMLElement, items: ISearchItem[], onClose: () => void) {
        // Render the summary
        new Dashboard({
            el,
            navigation: {
                title: "External Shares",
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

                    // Order by the 1st column by default; ascending
                    dtProps.order = [[0, "asc"]];

                    // Return the properties
                    return dtProps;
                },
                columns: [
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
                        name: "FileExtension",
                        title: "File Extension"
                    },
                    {
                        name: "ViewsLifeTime",
                        title: "Views"
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

        // Set the query
        let query: Types.Microsoft.Office.Server.Search.REST.SearchRequest = {
            Querytext: `ViewableByExternalUsers: true IsDocument: true path: ${DataSource.SiteContext.SiteFullUrl}`,
            RowLimit: 500,
            SelectProperties: {
                results: [
                    "Author", "FileExtension", "LastModifiedTime", "ListId",
                    "Path", "SPSiteUrl", "SPWebUrl", "Title", "WebId", "ViewsLifeTime"
                ]
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

            // Render the summary
            this.renderSummary(el, search.results, onClose);

            // Hide the loading dialog
            LoadingDialog.hide();
        });
    }
}