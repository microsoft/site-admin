import { Dashboard, Documents, LoadingDialog } from "dattatable";
import { Components, ContextInfo, Search, Types, Web } from "gd-sprest-bs";
import { fileEarmarkText } from "gd-sprest-bs/build/icons/svgs/fileEarmarkText";
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
    "Author", "FileExtension", "HitHighlightedSummary", "LastModifiedTime",
    "ListId", "Path", "SPSiteUrl", "SPWebUrl", "Title", "WebId"
]

// The valid file extensions for regex patterns
const FileExtensions = ["docx", "xlsx", "pptx", "pdf", "txt", "csv"];

export class SearchDocs {
    private static _loadOneDrive: boolean = false;

    // Analyzes the libraries of a site
    private static analyzeLibraries(url: string, regExPattern: RegExp) {
    }

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
    static getFormFields(fileExt: string = "", keywords: string = ""): Components.IFormControlProps[] {
        let ctrlRegEx: Components.IFormControl;
        let ctrlSearchTerms: Components.IFormControl;
        let ctrlSearchType: Components.IFormControl;
        return [
            {
                label: "Search Type",
                name: "SearchType",
                className: "mb-3",
                type: Components.FormControlTypes.Dropdown,
                required: true,
                items: [
                    { text: "Keyword", value: "Keyword", isSelected: true },
                    { text: "RegEx Pattern", value: "RegExPattern" }
                ],
                onControlRendered: ctrl => { ctrlSearchType = ctrl; },
                onChange: (item) => {
                    // See which one is selected
                    if (item.value == "Keyword") {
                        // Show/Hide the controls
                        ctrlRegEx.hide();
                        ctrlSearchTerms.show();
                    } else {
                        // Show/Hide the controls
                        ctrlRegEx.show();
                        ctrlSearchTerms.hide();
                    }
                }
            } as Components.IFormControlPropsDropdown,
            {
                label: "Search Terms",
                name: "SearchTerms",
                className: "mb-3",
                description: "Enter the search terms using quotes for phrases [Ex: movie \"social media\" show]",
                type: Components.FormControlTypes.TextField,
                required: true,
                value: keywords,
                errorMessage: "You must enter at least 1 search term.",
                onControlRendered: ctrl => { ctrlSearchTerms = ctrl; }
            },
            {
                label: "RegEx Pattern",
                name: "RegExPattern",
                className: "mb-3 d-none",
                description: "Enter the regular expression pattern to search for.",
                type: Components.FormControlTypes.TextField,
                required: true,
                value: keywords,
                errorMessage: "You must enter at least 1 search term.",
                onControlRendered: ctrl => { ctrlRegEx = ctrl; },
                onValidate: (ctrl, results) => {
                    // Ensure a pattern exists
                    if ((results.value || "").trim().length == 0) {
                        // Invalidate the extensions
                        results.isValid = false;
                        results.invalidMessage = "A regex pattern is required.";
                    } else {
                        // Validate the regex pattern
                        try {
                            // Create a new regex
                            new RegExp(results.value);
                        } catch (ex) {
                            // Invalidate the regex pattern
                            results.isValid = false;
                            results.invalidMessage = "The regex pattern is not valid.";
                        }
                    }

                    // Return the results
                    return results;
                }
            },
            {
                label: "File Types",
                name: "FileTypes",
                className: "mb-3",
                type: Components.FormControlTypes.TextField,
                value: fileExt,
                onValidate: (ctrl, results) => {
                    // See if a regex pattern was entered
                    if (ctrlSearchType.getValue().value == "RegExPattern" && results.value) {
                        // Validate the extension
                        let exts: string[] = results.value.split(' ');
                        for (let i = 0; i < exts.length; i++) {
                            let ext = exts[i].trim().toLowerCase();
                            if (ext) {
                                // See if it's a valid extension
                                if (FileExtensions.indexOf(ext) < 0) {
                                    // Invalidate the extensions
                                    results.isValid = false;
                                    results.invalidMessage = "The file extension '" + ext +
                                        "' is not valid. Valid extensions are: " + FileExtensions.join(", ") + ".";
                                }
                            }
                        }
                    }

                    // Return the results
                    return results;
                }
            }
        ];
    }

    // Renders the search summary
    private static renderSummary(el: HTMLElement, auditOnly: boolean, items: ISearchItem[], onClose: () => void) {
        // Render the summary
        new Dashboard({
            el,
            navigation: {
                title: "Search Documents",
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
                        name: "HitHighlightedSummary",
                        title: "Search Result",
                        onRenderCell: (el, col, item: ISearchItem) => {
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

                            // Clear the element
                            el.innerHTML = "";

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
                            let tooltips = Components.TooltipGroup({
                                el,
                                tooltips: [
                                    {
                                        content: "Click to view the document.",
                                        btnProps: {
                                            className: "pe-2 py-1",
                                            iconClassName: "mx-1",
                                            iconType: fileEarmarkText,
                                            iconSize: 24,
                                            text: "View",
                                            type: Components.ButtonTypes.OutlinePrimary,
                                            onClick: () => {
                                                // View the file
                                                window.open(Documents.isWopi(`${item.Title}.${item.FileExtension}`) ? item.SPWebUrl + "/_layouts/15/WopiFrame.aspx?sourcedoc=" + item.Path + "&action=view" : item.Path, "_blank");
                                            }
                                        }
                                    },
                                    {
                                        content: "Click to download the document.",
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
                                    }
                                ]
                            });

                            // Add the option to delete
                            if (!auditOnly) {
                                tooltips.add({
                                    content: "Click to delete the document.",
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
                                });
                            }
                        }
                    }
                ]
            }
        });
    }

    // Runs the report
    static run(el: HTMLElement, auditOnly: boolean, values: { [key: string]: string }, onClose: () => void) {
        this._loadOneDrive = values["LoadOneDrive"] == "true";

        // Show a loading dialog
        LoadingDialog.setHeader("Searching Site");
        LoadingDialog.setBody("Searching the content on this site...");
        LoadingDialog.show();

        // Get the form values
        let fileExt = values["FileTypes"] ? values["FileTypes"].split(' ') : null;
        let regExPattern = values["RegExPattern"] ? new RegExp(values["RegExPattern"]) : null;
        let searchTerms = (values["SearchTerms"] || "").split(' ');
        let searchType = (values["SearchType"] || "");

        // Set the query
        let query: Types.Microsoft.Office.Server.Search.REST.SearchRequest = {
            Querytext: `${searchTerms.join(" OR ")} IsDocument: true path: ${this._loadOneDrive ? DataSource.OneDriveWeb.Url : DataSource.SiteContext.SiteFullUrl}`,
            RowLimit: 500,
            SelectProperties: {
                results: [
                    "Author", "FileExtension", "HitHighlightedSummary", "LastModifiedTime",
                    "ListId", "Path", "SPSiteUrl", "SPWebUrl", "Title", "WebId"
                ]
            }
        };

        // See if file extensions exist
        if (fileExt) {
            // Set the filter
            query.RefinementFilters = {
                results: [`fileExtension:or("${fileExt.join('", "')}")`]
            };
        }

        // Set the url
        let url = this._loadOneDrive ? ContextInfo.siteAbsoluteUrl : DataSource.SiteContext.SiteFullUrl;

        // See if we are doing a keyword search
        if (searchType == "Keyword") {
            // Search for the content
            Search.postQuery({
                getAllItems: true,
                url,
                targetInfo: { requestDigest: this._loadOneDrive ? ContextInfo.formDigestValue : DataSource.SiteContext.FormDigestValue },
                query
            }).then(search => {
                // Clear the element
                while (el.firstChild) { el.removeChild(el.firstChild); }

                // Render the summary
                this.renderSummary(el, auditOnly, search.results, onClose);

                // Hide the loading dialog
                LoadingDialog.hide();
            });
        } else {
            // Search the libraries
            this.analyzeLibraries(url, regExPattern)
        }
    }
}