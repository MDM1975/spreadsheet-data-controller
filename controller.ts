/**
 * @function main - This is the entry point of the script. It accepts an Excel workbook and a CSV string as input.
 * @param workbook - The Excel workbook to be updated.
 * @param csv - The CSV string to be parsed.
 * @param key_column - The name of the column to be used as the key.
 */
function main(workbook: ExcelScript.Workbook, csv: string, key_column: string): void {
    /** @type {Parser.ParsedCollection} csv_data - The parsed CSV data. */
    const csv_data: Parser.ParsedCollection = Parser.parseCSV(csv, key_column);

    /**
     * @type {Parser.ParsedCollection} excel_data - The parsed Excel data. 
     * @type {Set<string>} excel_data.columns - The column names of the Excel data. 
     */
    const excel_data: { columns: Set<string>; rows: Parser.ParsedCollection; } = Parser.parseExcel(workbook, key_column);

    /** @type {SpreadsheetController.ControlerData} controller_data - The data to be passed to the controller. */
    const controller_data: SpreadsheetController.ControlerData = {
        csv_data,
        excel_data: excel_data.rows,
        excel_columns: excel_data.columns
    };

    /**
     * @function SpreadsheetController.run - The main function of the SpreadsheetController namespace.
     * @param {SpreadsheetController.ControlerData} controller_data - The data to be passed to the controller.
     * @param {ExcelScript.Workbook} workbook - The Excel workbook to be updated.
     */
    SpreadsheetController.run(controller_data, workbook);
}

/**
 * @namespace Parser - The namespace containing the functions for parsing CSV and Excel data.
 * @description The Parser namespace contains the functions for parsing CSV data (parseCSV()) and Excel data (parseExcel()). The parseCSV() function processes a CSV string and returns a collection of parsed data, while the parseExcel() function processes the Excel workbook and returns a collection of parsed data along with the column names.
 */
namespace SpreadsheetController {
    /**
     * @interface ControlerData - The data to be passed to the controller.
     * @property {Parser.ParsedCollection} csv_data - The parsed CSV data.
     * @property {Parser.ParsedCollection} excel_data - The parsed Excel data.
     * @property {Set<string>} excel_columns - The column names of the Excel data.
     */
    export interface ControlerData {
        csv_data: Parser.ParsedCollection;
        excel_data: Parser.ParsedCollection;
        excel_columns: Set<string>;
    }

    /**
     * @interface SheetRow - The data to be passed to the controller.
     * @property {number} row_index - The index of the row to be updated.
     * @property {SheetCell[]} cells - The cells to be updated.
     */
    interface SheetRow {
        row_index: number;
        cells: {
            cell_index: number;
            value: string;
        }[];
    }

    /**
     * @function run - The main function of the SpreadsheetController namespace.
     * @param {ControlerData} data - The data to be passed to the controller.
     * @param {ExcelScript.Workbook} workbook - The Excel workbook to be updated.
     * @returns {void}
     */
    export function run(data: ControlerData, workbook: ExcelScript.Workbook): void {
        /** @type {Parser.ParsedCollection} csvData - The parsed CSV data. */
        const csvData: Parser.ParsedCollection = { ...data.csv_data };

        /** @type {Parser.ParsedCollection} excelData - The parsed Excel data. */
        const excelData: Parser.ParsedCollection = { ...data.excel_data };

        /** @type {Map<string, number>} excelColumns - The column names of the Excel data. */
        const excelColumns: Map<string, number> = new Map(Array.from(data.excel_columns).map((column, index) => [column, index]));

        /** @type {SheetRow[]} rowsToAdd - The rows to be added to the worksheet. */
        const { rowsToAdd, rowsToUpdate } = partition(csvData, excelData, excelColumns);

        if (rowsToAdd.length > 0 || rowsToUpdate.length > 0) {
            updateWorksheet(workbook, rowsToAdd, rowsToUpdate);
        }
    }

    /**
     * @function partition - Partitions the data into rows to be added and rows to be updated.
     * @param {Parser.ParsedCollection} csvData - The parsed CSV data.
     * @param {Parser.ParsedCollection} excelData - The parsed Excel data.
     * @param {Map<string, number>} excelColumns - The column names of the Excel data.
     * @returns {{ rowsToAdd: SheetRow[]; rowsToUpdate: SheetRow[] }}
     */
    function partition(
        csvData: Parser.ParsedCollection,
        excelData: Parser.ParsedCollection,
        excelColumns: Map<string, number>
    ): { rowsToAdd: SheetRow[]; rowsToUpdate: SheetRow[] } {
        /** @type {SheetRow[]} rowsToAdd - The rows to be added to the worksheet. */
        const rowsToAdd: SheetRow[] = [];

        /** @type {SheetRow[]} rowsToUpdate - The rows to be updated on the worksheet. */
        const rowsToUpdate: SheetRow[] = [];

        for (const [key, value] of Object.entries(csvData)) {
            if (excelData[key]) {
                /** @type {number} row_index - The index of the row to be updated. */
                const row_index: number = excelData[key].row ?? 0;

                /** @type {SheetCell[]} cells - The cells to be updated. */
                const cells: SheetRow["cells"] = [];

                for (const cell of value.cells) {
                    /** @type {Parser.ParsedCell} excel_cell - The cell from the Excel data. */
                    const excel_cell: Parser.ParsedData | undefined = excelData[key].cells.find((excel_cell) => excel_cell.column === cell.column);

                    if (excel_cell && excel_cell.value !== cell.value) {
                        /** @type {number} cell_index - The index of the cell to be updated. */
                        const cell_index: number | undefined = excelColumns.get(cell.column);
                        if (cell_index !== undefined) {
                            cells.push({ cell_index, value: cell.value });
                        }
                    }
                }

                rowsToUpdate.push({ row_index, cells });
            } else {
                /** @type {number} row_index - The index of the row to be added. */
                const row_index: number = Object.keys(excelData).length + rowsToAdd.length + 1;

                /** @type {SheetCell[]} cells - The cells to be added. */
                const cells: SheetRow["cells"] = [];

                for (const cell of value.cells) {
                    /** @type {number} cell_index - The index of the cell to be added. */
                    const cell_index: number | undefined = excelColumns.get(cell.column);
                    if (cell_index !== undefined) {
                        cells.push({ cell_index, value: cell.value });
                    }
                }

                rowsToAdd.push({ row_index, cells });
            }
        }

        return { rowsToAdd, rowsToUpdate };
    }

    /**
     * @function updateWorksheet - Updates the worksheet with the provided data.
     * @param {ExcelScript.Workbook} workbook - The Excel workbook to be updated.
     * @param {SheetRow[]} rowsToAdd - The rows to be added to the worksheet.
     * @param {SheetRow[]} rowsToUpdate - The rows to be updated on the worksheet.
     * @returns {void}
     */
    function updateWorksheet(
        workbook: ExcelScript.Workbook,
        rowsToAdd: SheetRow[],
        rowsToUpdate: SheetRow[]
    ): void {
        /** @type {ExcelScript.Worksheet} worksheet - The worksheet to be updated. */
        const worksheet: Exclude.Workbook = workbook.getActiveWorksheet();

        for (const row of rowsToAdd.concat(rowsToUpdate)) {
            for (const cell of row.cells) {
                worksheet.getCell(row.row_index, cell.cell_index).setValue(cell.value);
            }
        }
    }
}

/**
 * @namespace Parser - The namespace for the parser functions.
 * @description The parser functions are used to parse the CSV and Excel data.
 */
namespace Parser {
    /**
     * @interface ParsedData - The data for a single cell.
     * @property {string} column - The column name of the cell.
     * @property {string} value - The value of the cell.
     */
    export interface ParsedData {
        column: string;
        value: string;
    }

    /**
     * @interface ParsedCollection - The data for a single row.
     * @property {string} key - The key of the row.
     * @property {number} row - The row number of the row.
     * @property {ParsedData[]} cells - The cells of the row.
     */
    export interface ParsedCollection {
        [key: string]: {
            row?: number;
            cells: ParsedData[];
        };
    }

    /**
     * @function parseCSV - Parses the CSV data.
     * @param {string} csv - The CSV data.
     * @param {string} key - The column name of the key.
     * @returns {ParsedCollection}
     */
    export function parseCSV(csv: string, key: string): ParsedCollection {
        /** @type {string[]} columns - The column names of the CSV data. */
        const [columns, ...rows]: string[][] = csv.split(/[\r\n]+/).map((line) => line.replace(/["\t]/g, '').trim().split(','));

        /** @type {ParsedCollection} data - The parsed CSV data. */
        const data: ParsedCollection = {};

        /** @type {number} id_index - The index of the Key column. */
        const id_index: number = columns.indexOf(key);

        for (const row of rows) {
            if (!row[id_index]) continue;

            /** @type {string} key - The key of the row. */
            const key: string = row[id_index];
            data[key] = { cells: [] }

            for (const [cell_index, cell] of Array.from(row.entries())) {
                data[key].cells.push({ column: columns[cell_index], value: Serialize.execute(cell) });
            }
        }

        return data;
    }

    /**
     * @function parseExcel - Parses the Excel data.
     * @param {ExcelScript.Workbook} workbook - The Excel workbook to be parsed.
     * @param {string} key - The column name of the key.
     * @returns {Set<string>, ParsedCollection} - The columns and rows of the Excel data.
     */
    export function parseExcel(workbook: ExcelScript.Workbook, key: string): { columns: Set<string>; rows: ParsedCollection; } {
        /** @type {string[]} columns - The column names of the Excel data. */
        const [columns, ...rows]: string[][] = workbook.getActiveWorksheet().getUsedRange().getValues().map(
            (row: (string | number | boolean)[]) => row.map((cell) => `${cell ?? ''}`.trim().toString())
        );

        /** @type {ParsedCollection} data - The parsed Excel data. */
        const data: ParsedCollection = {};

        /** @type {number} id_index - The index of the Key column. */
        const id_index: number = columns.indexOf(key);

        for (const [row_index, row] of Array.from(rows.entries())) {
            /** @type {string} key - The key of the row. */
            const key: string = row[id_index];
            data[key] = { row: row_index + 1, cells: [] }
            for (const [cell_index, cell] of Array.from(row.entries())) {
                data[key].cells.push({ column: columns[cell_index], value: cell });
            }
        }

        return { columns: new Set(columns), rows: data };
    }
}

/**
 * @namespace Serialize - The namespace for the serialize functions.
 * @description The serialize functions are used to serialize the data for the CSV and Excel files.
 */
namespace Serialize {
    /** @constant {Set<string>} BOOLEAN_VALUE_TEST - The set of values that should be treated as booleans. */
    const BOOLEAN_VALUE_TEST = new Set(['TRUE', 'FALSE', 'YES', 'NO', 'Y', 'N', 'T', 'F']);

    /** @constant {RegExp} DATE_VALUE_TEST - The regular expression used to test if a value is a date. */
    const DATE_VALUE_TEST = new RegExp(/^\d{1,2}\/\d{1,2}\/\d{2,4}$/);

    /** @constant {Set<string>} TRUTHY_VALUES - The set of values that should be treated as truthy. */
    const TRUTHY_VALUES = new Set(['TRUE', 'YES', 'Y', 'T']);

    /** @constant {number} MS_PER_DAY - The number of milliseconds in a day. */
    const MS_PER_DAY = 24 * 60 * 60 * 1000;

    /** @constant {number} EXCEL_EPOCH_OFFSET - The number of days between the Excel epoch and the Unix epoch. */
    const EXCEL_EPOCH_OFFSET = 25569;

    /**
     * @function execute - Serializes the value for the CSV and Excel files.
     * @param {string} value - The value to be serialized.
     * @returns {string}
     */
    export function execute(value: string): string {
        return isDate(value) ? serializeDate(value) : isBoolean(value) ? serializeBoolean(value) : value;
    }

    /**
     * @function isDate - Tests if the value is a date.
     * @param {string} value - The value to be tested.
     * @returns {boolean}
     */
    function isDate(value: string): boolean {
        return DATE_VALUE_TEST.test(value);
    }

    /**
     * @function isBoolean - Tests if the value is a boolean.
     * @param {string} value - The value to be tested.
     * @returns {boolean}
     */
    function isBoolean(value: string): boolean {
        return BOOLEAN_VALUE_TEST.has(value.toUpperCase());
    }

    /**
     * @function serializeBoolean - Serializes the boolean value for the CSV and Excel files.
     * @param {string} value - The value to be serialized.
     * @returns {string}
     */
    function serializeBoolean(value: string): string {
        return TRUTHY_VALUES.has(value.toUpperCase()) ? 'true' : 'false';
    }

    /**
     * @function serializeDate - Serializes the date value for the CSV and Excel files.
     * @param {string} value - The value to be serialized.
     * @returns {string}
     */
    function serializeDate(value: string): string {
        return (Math.floor(new Date(value).getTime() / MS_PER_DAY) + EXCEL_EPOCH_OFFSET).toString();
    }
}




