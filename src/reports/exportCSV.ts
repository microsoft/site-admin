/**
 * Exports data to a csv
 */
export class ExportCSV {
    private _columns: string[] = null;
    private _filename: string = null;
    private _rows: any[] = null;

    // Constructor
    constructor(filename: string, columns: string[], rows: any[]) {
        // Save the properties
        this._columns = columns;
        this._rows = rows;
        this._filename = filename;

        // Create CSV
        this.generateCSV();
    }

    // Generates the csv
    private generateCSV() {
        let csv = "";

        // Set the header row
        csv = this._columns.join(',') + '\n';

        // Parse the rows
        for (let i = 0; i < this._rows.length; i++) {
            let data = this._rows[i];
            let row = [];

            // Parse the columns
            for (let i = 0; i < this._columns.length; i++) {
                let col = this._columns[i];

                // Add the column value
                row.push(data[col] || "");
            }

            // Add the row to the csv
            csv += '"' + row.join('","') + '"\n';
        }

        // See if this is IE or Mozilla
        if (Blob && navigator && navigator["msSaveBlob"]) {
            // Download the file
            navigator["msSaveBlob"](new Blob([csv], { type: "data:text/csv;charset=utf-8;" }), this._filename);
        } else {
            // Generate an anchor
            var anchor = document.createElement("a");
            anchor.download = this._filename;
            anchor.href = "data:text/csv;charset=utf-8," + encodeURIComponent(csv);
            anchor.target = "__blank";

            // Download the file
            anchor.click();
        }
    }
}