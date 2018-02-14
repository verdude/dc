class WBContainer {
    constructor(wb) {
        this.error = false;
        this.filename = null;
        this.errorMessage = null;

        wb = JSON.parse(wb);
        this.sheetnames = Object.keys(wb);
        if (this.sheetnames.length > 1) {
            this.error = true;
            this.errorMessage = "Multiple sheets found in WorkBook.";
            return;
        } else if (this.sheetnames.length === 0) {
            this.error = true;
            this.errorMessage = "No sheets found in WorkBook";
            return;
        }

        this.data = wb[this.sheetnames[0]];
        this.title = this.sheetnames[0];
        this.numRows = this.data.length;
        this.numColumns = this.data[0].length;
        this.loaded = true;
    }

    ready() {
        return !this.error && this.loaded;
    }

    getCellValue(row, column) {
        return this.data[row][column];
    }
}
