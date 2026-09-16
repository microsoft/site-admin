import { CanvasForm } from "dattatable";
import { Components } from "gd-sprest-bs";

/**
 * Skipped Lists Dialog
 */
export class SkippedListsDialog {
    constructor(lists: string[]) {
        // Show the dialog
        this.showDialog(lists);
    }

    // Shows the dialog
    private showDialog(lists: string[]) {
        // Initialize the form
        CanvasForm.setHeader("Skipped Lists");
        CanvasForm.setSize(Components.OffcanvasSize.Medium2);
        CanvasForm.setType(Components.OffcanvasTypes.End);

        // Parse the lists to display
        let items: Components.IListGroupItem[] = [];
        lists.forEach(list => {
            items.push({ content: list });
        });

        // Render the list
        Components.ListGroup({
            el: CanvasForm.BodyElement,
            items
        });

        // Show the form
        CanvasForm.show();
    }
}