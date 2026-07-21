import { CanvasForm, Dashboard, Modal } from "dattatable";
import { Components, Helper } from "gd-sprest-bs";

export interface IRegexPattern {
    title: string;
    patterns: string[];
}

interface IRegexPatternImport extends IRegexPattern {
    errors: string[];
}

/**
 * Dialog to map regex patterns to a category for a dropdown list.
 */
export class RegexDialog {
    private _dashboard: Dashboard = null;
    private _patterns: IRegexPattern[] = [];
    private _onUpdate: (patterns: string) => void;

    // Constructor
    constructor(patterns: string, onUpdate: (patterns: string) => void) {
        // Set the update callback
        this._onUpdate = onUpdate;

        // Initializes the dialog
        this.init(patterns);
    }

    // Initializes the dialog
    private init(patterns: string) {
        // Clear the modal
        Modal.clear();
        Modal.setType(Components.ModalTypes.Large);

        // Set the header
        Modal.setHeader("Regex Patterns");

        // Render the table
        this.renderTable(patterns);

        // Set the footer
        Components.TooltipGroup({
            el: Modal.FooterElement,
            tooltips: [
                {
                    content: "Click to update the regex patterns.",
                    btnProps: {
                        text: "Update",
                        type: Components.ButtonTypes.OutlinePrimary,
                        onClick: () => {
                            // Get the patterns
                            const patterns = JSON.stringify(this._patterns);

                            // Call the event
                            this._onUpdate(patterns);

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

    // Render the table
    private renderTable(patterns: string) {
        // Try to get the patterns
        try {
            // Parse the patterns
            this._patterns = JSON.parse(patterns);
        } catch {
            this._patterns = [];
        }

        // Render the dashboard
        this._dashboard = new Dashboard({
            el: Modal.BodyElement,
            navigation: {
                showFilter: false,
                items: [
                    {
                        className: "btn-outline-light me-2",
                        isButton: true,
                        text: "Add Category",
                        onClick: () => {
                            // Show the form
                            this.showForm(null, item => {
                                // Append the item to the patterns
                                this._patterns.push(item);

                                // Add the row to the dashboard
                                this._dashboard.Datatable.addRow(item);
                            });
                        }
                    },
                    {
                        className: "btn-outline-light me-2",
                        isButton: true,
                        text: "Clear All",
                        onClick: () => {
                            // Clear the patterns
                            this._patterns = [];

                            // Update the dashboard
                            this._dashboard.refresh(this._patterns);
                        }
                    },
                    {
                        className: "btn-outline-light",
                        isButton: true,
                        text: "Import",
                        onClick: () => {
                            // Show the form
                            this.showImportForm(items => {
                                // Append the item to the patterns
                                this._patterns = this._patterns.concat(items).sort((a, b) => {
                                    if (a.title < b.title) return -1;
                                    if (a.title > b.title) return 1;
                                    return 0;
                                });

                                // Add the row to the dashboard
                                this._dashboard.refresh(this._patterns);
                            });
                        }
                    }
                ]
            },
            table: {
                onRendering: (dtProps) => {
                    dtProps.columnDefs = [
                        {
                            "targets": 2,
                            "orderable": false,
                            "searchable": false
                        }
                    ];
                },
                rows: this._patterns,
                columns: [
                    {
                        name: "title",
                        title: "Category"
                    },
                    {
                        name: "patterns",
                        title: "Patterns",
                        onRenderCell: (el, col, item: IRegexPattern) => {
                            // Set the style to allow horizontal scrolling
                            el.style.overflowX = "auto";
                            el.style.maxWidth = "500px";

                            // Show the patterns
                            el.innerHTML = item.patterns.join("<br/>");
                        }
                    },
                    {
                        name: "",
                        title: "",
                        onRenderCell: (el, col, item: IRegexPattern, rowIdx) => {
                            // Render a button
                            Components.TooltipGroup({
                                el,
                                tooltips: [
                                    {
                                        content: "Click to modify the regex patterns.",
                                        btnProps: {
                                            text: "Edit",
                                            onClick: () => {
                                                // Show a button to edit the pattern
                                                this.showForm(item, newItem => {
                                                    // Update the item
                                                    for (let i = 0; i < this._patterns.length; i++) {
                                                        let pattern = this._patterns[i];
                                                        if (pattern.title == item.title) {
                                                            // Update the item
                                                            this._patterns[i] = newItem;
                                                            break;
                                                        }
                                                    }

                                                    // Update the table
                                                    this._dashboard.refresh(this._patterns);
                                                });
                                            }
                                        }
                                    },
                                    {
                                        content: "Click to remove this regex category.",
                                        btnProps: {
                                            text: "Delete",
                                            type: Components.ButtonTypes.OutlineDanger,
                                            onClick: () => {
                                                // Remove the row
                                                this._dashboard.Datatable.datatable.row(rowIdx).remove().draw(false);
                                                // Parse the patterns
                                                for (let i = 0; i < this._patterns.length; i++) {
                                                    if (this._patterns[i].title == item.title) {
                                                        // Remove this item
                                                        this._patterns.splice(i, 1);
                                                        break;
                                                    }
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

    // Show the form
    private showImportForm(onImportItems: (items: IRegexPattern[]) => void) {
        // Clear the canvas form
        CanvasForm.clear();
        CanvasForm.setSize(Components.OffcanvasSize.Medium2);

        // Set the header
        CanvasForm.setHeader("Regex Pattern By JSON");

        // Set the body
        CanvasForm.BodyElement.innerHTML = `
            <div class="mb-2">
                Click the "Download Template" button to download a CSV template. The first column will be the category with each regex pattern associated with the category separated by a comma.
            </div>
        `;

        // Set the body
        let importItems: IRegexPatternImport[] = [];
        let importDashboard = new Dashboard({
            el: CanvasForm.BodyElement,
            navigation: {
                items: [
                    {
                        className: "btn-outline-light",
                        isButton: true,
                        text: "Download Template",
                        onClick: () => {
                            // Create the template
                            const csv = "category,pattern1,pattern2,pattern3,pattern4,pattern5,pattern6,pattern7,pattern8,pattern9,pattern10\r\n";
                            const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
                            const url = URL.createObjectURL(blob);

                            // Send the template to the user
                            const link = document.createElement("a");
                            link.href = url;
                            link.download = "regex-template.csv";
                            link.click();
                            URL.revokeObjectURL(url);
                        }
                    },
                    {
                        className: "btn-outline-light ms-2",
                        isButton: true,
                        text: "Import Template",
                        onClick: () => {
                            // Show a file upload dialog
                            Helper.ListForm.showFileDialog(["csv"]).then(file => {
                                // Read the csv file
                                file.src.text().then(data => {
                                    let rows = data.split("\r\n");

                                    // Clear the rows
                                    importItems = [];

                                    // Parse the rows to import
                                    for (let i = 1; i < rows.length; i++) {
                                        let row = rows[i].trim();
                                        let data = row.split(",");
                                        if (data.length > 1) {
                                            // Parse the patterns
                                            let item: IRegexPatternImport = {
                                                title: data[0].trim(),
                                                patterns: [],
                                                errors: []
                                            };

                                            // Parse the patterns
                                            for (let j = 1; j < data.length; j++) {
                                                let isValid = true;

                                                // Add the pattern
                                                let pattern = data[j].trim();
                                                if (pattern) {
                                                    // Try to parse the value
                                                    try { new RegExp(pattern); }
                                                    catch {
                                                        // Set the flag
                                                        isValid = false;
                                                    }

                                                    // See if it's valid
                                                    if (isValid) {
                                                        // Add the pattern
                                                        item.patterns.push(pattern);
                                                    } else {
                                                        // Add the pattern
                                                        item.errors.push(pattern);
                                                    }
                                                }
                                            }

                                            // Append the item
                                            importItems.push(item);
                                        }
                                    }

                                    // Update the dashboard
                                    importDashboard.refresh(importItems);
                                });
                            });
                        }
                    }
                ]
            },
            table: {
                rows: [],
                columns: [
                    {
                        name: "title",
                        title: "Category"
                    },
                    {
                        name: "patterns",
                        title: "Patterns",
                        onRenderCell: (el, col, item) => {
                            // Set the style to allow horizontal scrolling
                            el.style.overflowX = "auto";
                            el.style.maxWidth = "500px";

                            // Show the patterns
                            el.innerHTML = item.patterns.join("<br/>");
                        }
                    },
                    {
                        name: "",
                        title: "Errors",
                        onRenderCell: (el, col, item) => {
                            // Set the style to allow horizontal scrolling
                            el.style.overflowX = "auto";
                            el.style.maxWidth = "500px";

                            // Show the errors
                            el.innerHTML = item.errors.join("<br/>");
                        }
                    }
                ]
            }
        });

        // Add a footer
        let elFooter = document.createElement("div");
        elFooter.classList.add("d-flex", "justify-content-end");
        CanvasForm.BodyElement.appendChild(elFooter);
        Components.ButtonGroup({
            el: elFooter,
            className: "mt-2",
            buttons: [
                {
                    text: "Import",
                    type: Components.ButtonTypes.OutlinePrimary,
                    onClick: () => {
                        // Call the event
                        onImportItems(importItems);

                        // Close the canvas form
                        CanvasForm.hide();
                    }
                },
                {
                    text: "Cancel",
                    type: Components.ButtonTypes.OutlineSecondary,
                    onClick: () => {
                        // Close the canvas form
                        CanvasForm.hide();
                    }
                }
            ]
        });

        // Show the form
        CanvasForm.show();
    }

    // Show the form
    private showForm(item: IRegexPattern, onUpdate: (item: IRegexPattern) => void) {
        // Clear the canvas form
        CanvasForm.clear();
        CanvasForm.setSize(Components.OffcanvasSize.Medium2);
        CanvasForm.setAutoClose(false);

        // Set the header
        CanvasForm.setHeader("Regex Pattern");

        // Set the body
        let form = Components.Form({
            el: CanvasForm.BodyElement,
            controls: [
                {
                    name: "title",
                    label: "Category",
                    type: Components.FormControlTypes.TextField,
                    required: true,
                    value: item?.title
                },
                {
                    name: "patterns",
                    label: "Regex Patterns",
                    description: "Enter the regex patterns for this category, one per line.",
                    type: Components.FormControlTypes.TextArea,
                    required: true,
                    rows: 6,
                    value: item ? item.patterns.join('\n') : ""
                } as Components.IFormControlPropsTextField,
            ]
        });

        // Add a footer
        let elFooter = document.createElement("div");
        elFooter.classList.add("d-flex", "justify-content-end");
        CanvasForm.BodyElement.appendChild(elFooter);
        Components.ButtonGroup({
            el: elFooter,
            className: "mt-2",
            buttons: [
                {
                    text: "Save",
                    type: Components.ButtonTypes.OutlinePrimary,
                    onClick: () => {
                        // Validate the form
                        if (form.isValid()) {
                            // Set the item
                            let values = form.getValues();

                            // Call the update event
                            onUpdate({
                                title: values["title"],
                                patterns: values["patterns"].split("\n").map(s => s.trim()).filter(s => s != "")
                            });

                            // Close the form
                            CanvasForm.hide();
                        }
                    }
                },
                {
                    text: "Cancel",
                    type: Components.ButtonTypes.OutlineSecondary,
                    onClick: () => {
                        // Close the canvas form
                        CanvasForm.hide();
                    }
                }
            ]
        });

        // Show the form
        CanvasForm.show();
    }
}