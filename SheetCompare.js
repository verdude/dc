class SheetCompare {
    constructor() {
        this.sheets = [];
        this.sheetTotal = 0;
        this.error = false;
        this.differences = [];
        this.alphabet = Array(26);
        for (let i = 0; i < this.alphabet.length; ++i) {
            this.alphabet[i] = String.fromCharCode(i + 'A'.charCodeAt(0));
        }
    }

    columnNumberToLetters(num) {
        let remainder = num % 26;
        let quotient = Math.floor(num / 26);
        let firstLetter = quotient>0?this.alphabet[quotient-1]:this.alphabet[remainder];
        let secondLetter = quotient>0?this.alphabet[remainder]:'';
        return firstLetter+secondLetter;
    }

    configure(num, caseSensitive) {
        // reset
        this.sheetTotal = num;
        this.sheets = [];
        this.differences = [];
        this.caseSensitive = caseSensitive;
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
            console.log("NO FILES PROVIDED...EXITING COMPARE PROCEDURE.");
            let data = {
                titles: this.sheets.map((x)=>x.title),
                differences: [],
                error: "No Sheets submitted"
            }
            return data;
        }
        console.log("INITIATING COMPARE PROCEDURE...");
        const sheets_rows = this.sheets.map((x)=>x.numRows);
        const sheets_cols = this.sheets.map((x)=>x.numColumns);
        const rows = Math.min.apply(Math, sheets_rows);
        const cols = Math.min.apply(Math, sheets_cols);

        let row_diff_msg = sheets_rows.every((n) => n === sheets_rows[0]) ? [] : this.sheets.map(x => x.sheetnames + " has " + x.numRows + " rows");
        let col_diff_msg = sheets_cols.every((n) => n === sheets_cols[0]) ? [] : this.sheets.map(x => x.sheetnames + " has " + x.numCols + " columns");
        let warning_msg = row_diff_msg.concat(col_diff_msg).join(", ");
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
                    this.differences.push({cells: cells, y: y+1, x: this.columnNumberToLetters(x)});
                }
            }
        }
        console.log("COMPARE PROCEDURE TERMINATING...");
        let data = {
            titles: this.sheets.map((x)=>x.title),
            differences: this.differences.filter((diff)=> {
                // might need to pass over twice, once to check for 
                // null and once to check for undefined
                // if they are meaningfully different values in excel
                return !diff.cells.every((c)=>c===null||c===undefined);
            }),
            warning: warning_msg.length > 0 ? warning_msg : undefined
        }
        console.log(`FOUND ${data.differences.length} DIFFERENCES`);
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
