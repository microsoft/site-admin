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

    // Gets the form fields to display
    static getFormFields(years?: string): Components.IFormControlProps[] {
        // Set the default # of months to search for
        let numbOfMonths = parseInt(years) > 0 ? parseInt(years) * 12 : 36;

        return [{
            name: "SelectedDate",
            label: "Select Date",
            description: "The date to find content older than.",
            errorMessage: "A date is required to run the query.",
            type: Components.FormControlTypes.DateTime,
            required: true,
            value: moment(Date.now()).subtract(numbOfMonths, "months").toISOString()
        }];
    }

    // Renders the search summary
    private static renderSummary(el: HTMLElement, auditOnly: boolean, items: ISearchItem[], onClose: () => void) {
        // Render the summary
        new Dashboard({
            el,
            navigation: {
                title: "Doc Retention",
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
                        new ExportCSV("docRetention.csv", CSVFields, items);
                    }
                }]
            },
            table: {
                rows: items,
                onRendering: dtProps => {
                    dtProps.columnDefs = [
                        {
                            "targets": 4,
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
                            let tooltips = Components.TooltipGroup({
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
                                    }
                                ]
                            });

                            // Add the delete option
                            if (!auditOnly) {
                                tooltips.add({
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
                                })
                            }
                        }
                    }
                ]
            }
        });
    }

    // Runs the report
    static run(el: HTMLElement, auditOnly: boolean, values: { [key: string]: string }, onClose: () => void) {
        // Show a loading dialog
        LoadingDialog.setHeader("Searching Site");
        LoadingDialog.setBody("Searching the site for files...");
        LoadingDialog.show();

        // Get the start date
        let startDate = moment(values["SelectedDate"]).format("YYYY-MM-DD");

        // Search the site
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
            // Clear the element
            while (el.firstChild) { el.removeChild(el.firstChild); }

            // Render the summary
            this.renderSummary(el, auditOnly, search.results, onClose);

            // Hide the loading dialog
            LoadingDialog.hide();
        });
    }
}