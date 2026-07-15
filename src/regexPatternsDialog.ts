import { CanvasForm, Dashboard, Modal } from "dattatable";
import { Components } from "gd-sprest-bs";

export interface IRegexPattern {
    title: string;
    patterns: string[];
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
                        className: "btn-outline-light",
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
                    }
                ]
            },
            table: {
                rows: this._patterns,
                columns: [
                    {
                        name: "title",
                        title: "Category"
                    },
                    {
                        name: "",
                        title: "Patterns",
                        onRenderCell: (el, col, item: IRegexPattern) => {
                            // Show the patterns
                            el.innerHTML = item.patterns.join("<br/>");
                        }
                    },
                    {
                        name: "",
                        title: "",
                        onRenderCell: (el, col, item: IRegexPattern, rowIdx) => {
                            // Render a button
                            Components.Button({
                                el,
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

                                        // Update this row
                                        this._dashboard.updateRow(rowIdx, newItem);
                                    });
                                }
                            });
                        }
                    }
                ]
            }
        });
    }

    // Show the form
    private showForm(item: IRegexPattern, onUpdate: (item: IRegexPattern) => void) {
        // Clear the canvas form
        CanvasForm.clear();
        CanvasForm.setSize(Components.OffcanvasSize.Medium2);

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