import { CanvasForm } from "dattatable";
import { Components } from "gd-sprest-bs";

// Skipped List Info
export interface ISkippedList {
    title: string;
    webUrl: string;
}

/**
 * Skipped Lists Dialog
 */
export class SkippedListsDialog {
    constructor(lists: ISkippedList[]) {
        // Show the dialog
        this.showDialog(lists);
    }

    // Shows the dialog
    private showDialog(lists: ISkippedList[]) {
        // Initialize the form
        CanvasForm.setHeader("Skipped Lists");
        CanvasForm.setSize(Components.OffcanvasSize.Medium2);
        CanvasForm.setType(Components.OffcanvasTypes.End);

        // Parse the lists to display
        let items: Components.IListGroupItem[] = [];
        lists.forEach(list => {
            items.push({ content: `<b>List: </b>${list.title}<br/><b>Web: </b>${list.webUrl}` });
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