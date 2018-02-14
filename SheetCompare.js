class SheetCompare {
    contructor() {
        this.sheets = [];
        this.sheetTotal = 0;
        this.error = false;
        this.differences = [];
    }

    configure(num, caseSensitive, send_results_proc) {
        // reset
        this.sheetTotal = num;
        this.sheets = [];
        this.differences = [];
        this.caseSensitive = caseSensitive;
        this.send_results_proc = send_results_proc;
        return this;
    }

    ready() {
        return !this.error && this.sheets.length === this.sheetTotal;
    }

    verbose_loggage(rows, cols) {
        var metrics = this.sheets.map((sheet) => {return `Title: ${sheet.title} -- Rows: ${sheet.numRows} -- Columns: ${sheet.numColumns} -- Filename: ${sheet.filename}`}).join("\n");
        console.log(
`Number Of Sheets to Compare: ${this.sheetTotal}
Sheets Metrics:
${metrics}`);
    }

    compare() {
        if (this.sheets.length <= 1) {
            console.log("NO DIFFERENCES IN FILES...EXITING COMPARE PROCEDURE.");
            return;
        }
        console.log("INITIATING COMPARE PROCEDURE...");
        const rows = Math.min.apply(Math, this.sheets.map((x)=>x.numRows));
        const cols = Math.min.apply(Math, this.sheets.map((x)=>x.numColumns));

        this.verbose_loggage(rows, cols);

        for (let y = 0; y < rows; y++) {
            for (let x = 0; x < cols; x++) {

                let cells = this.sheets.map((sheet) => {
                    let cell = sheet.getCellValue(y, x);

                    if (typeof cell === "string") {
                        if (this.caseSensitive)
                            return cell.trim();
                        else
                            return cell.trim().toLowerCase();
                    } else { return cell; }
                });

                // check for equality
                if (!cells.reduce((a, b) => { return (a === b) ? a : NaN})) {
                    this.differences.push({cells: cells, y: y, x: x});
                }
            }
        }
        console.log("COMPARE PROCEDURE TERMINATING...");
        console.log(`FOUND ${this.differences.length} DIFFERENCES`);
        var data = {
            titles: this.sheets.map((x)=>x.title),
            differences: this.differences
        }
        return data;
    }

    registerSheet(sheet) {
        if (!(this.sheets instanceof Array)) {
            return;
        }
        if (sheet instanceof WBContainer) {
            if (sheet.ready()) {
                this.sheets.push(sheet);
            } else {
                 console.log("Error: " + sheet.errorMessage);
            }
        } else {
            console.log("Attempted to insert a " + (typeof sheet) +
                    " into the sheets Array. This should have been an" +
                    " instanceof Sheet");
        }
    }
}
