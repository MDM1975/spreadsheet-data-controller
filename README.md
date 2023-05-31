# Spreadsheet Data Controller Tool
This repository contains a tool developed to manage and manipulate data from CSV files and Excel workbooks. The tool provides functionality to parse CSV and Excel data, serialize date and boolean values, and partition data for efficient updates. The main component of this project is the data controller, implemented in TypeScript, which orchestrates all operations from data ingestion to partitioning. The data controller is designed as a custom connector for the Power Automate tool in the Power Platform environment to automate data processing and manipulation tasks in Excel workbooks.

## Key Responsibilities of the Controller
The data controller has the following key responsibilities:
* **Data Import:** The controller handles data import from CSV files and Excel workbooks. CSV data is converted into a JSON object using the Parser.parseCSV function, while Excel data is handled through the Parser.parseExcel function.
* **Data Serialization:** After data import, the controller serializes the data by standardizing date and boolean values into a format that can be processed further. This is achieved using the Serialize.execute function.
* **Data Partitioning:** One of the primary roles of the controller is to partition the CSV and Excel data, enabling efficient updates to the workbook. The SpreadsheetController.partition function is used for this purpose.

## Design Considerations and Performance Analysis
The script incorporates several design considerations to ensure optimal efficiency and performance, especially when dealing large amounts of data. Here are the critical design decisions and their impact on performance:

### **Utilizing Object and Set Data Structures**
The script leverages Object and Set data structures for efficient data storage, providing faster lookups and better memory management than arrays.
* **Object Data Structure:** Parsed data from the Excel workbook and the CSV file are stored in the Object data structure. This key-value pair storage enhances performance by enabling faster lookups and data manipulation.
* **Set Data Structure:** The Set data structure stores column names, ensuring uniqueness and facilitating faster lookups when indexing specific columns in the dataset. These data structures significantly improve performance and scalability, particularly when handling large volumes of data.

### **Partitioning CSV Data for Efficient Excel Workbook Updates**
The script implements a partitioning strategy to update the Excel workbook more efficiently. The number of required updates on the workbook is reduced by separating CSV data into rows to be added and updated.

Partitioning is achieved by comparing the data from the CSV file with the data from the Excel workbook. If a row exists in the workbook, it is added to the rowsToUpdate array if any corresponding cell values have been changed. Otherwise, the row is added to the rowsToAdd collection. This partitioning approach streamlines the update process, resulting in fewer updates required and improved overall performance.

### **Data Normalization and Scaling Considerations**
The script includes functions to normalize data, specifically addressing boolean and date values, to ensure proper formatting when updating the Excel workbook. Additionally, as the program scales, it is essential to consider potential limitations, such as batch sizing governor limits, and implement appropriate strategies to handle larger data sets.

### **Normalizing Boolean and Date Values**
To resolve issues with string date values from the CSV file not being recognized as date values in Excel, the script converts date values to serial numbers representing the days since January 1, 1900. This conversion allows Excel to identify and format the values correctly as dates.

Addressing Batch Sizing Governor Limits
As the volume of processed data increases, it's important to consider the governor limits established by the Power Automate platform. To mitigate this, batch processing can be implemented by dividing data into smaller segments and processing them sequentially or concurrently. This approach ensures that the program can scale up.

## Big-O Time Complexity Analysis
This script consists of three parts: parsing CSV data, Excel data, and partitioning data. The time complexity of each of these operations is analyzed below.

### **Parsing CSV Data $O(n^2)$:**
The time complexity of parsing CSV data is dominated by the nested loop that iterates through each row and each cell, resulting in a time complexity of O(n^2), where n is the total number of cells in the CSV data.

#### Time Complexity Analysis
* Splitting the CSV data into rows takes linear time, $O(n)$, where n is the number of rows in the CSV data.
* Creating an empty object takes constant time, $O(1)$.
* Retrieving the index of the key column takes linear time, $O(n)$, where n is the number of columns in the CSV data.
* Iterating through each row in the CSV data takes linear time, $O(n)$, where n is the number of rows in the CSV data.
* Skipping the row if the key column is empty takes constant time, $O(1)$.
* Retrieving the key value takes constant time, $O(1)$.
* Initializing an empty object takes constant time, $O(1)$.
* Iterating through each cell in the row takes linear time, $O(n)$, where n is the number of cells in the row.
* Adding the cell to the data object takes constant time, $O(1)$.

```typescript
function parseCSV(csv: string, key: string): ParsedCollection {
    // Splitting the CSV data into rows takes linear time -> O(n), where n is the number of rows in the CSV data.
    const [columns, ...rows]: string[][] = csv.split(/[\r\n]+/).map((line) => line.replace(/["\t]/g, '').trim().split(','));

    // Creating an empty object takes constant time -> O(1).
    const data: ParsedCollection = {};
    // Retrieving the index of the key column takes linear time -> O(n), where n is the number of columns in the CSV data.
    const id_index: number = columns.indexOf(key);

    // Iterating through each row in the CSV data takes linear time -> O(n), where n is the number of rows in the CSV data.
    for (const row of rows) {
        // Skipping the row if the key column is empty takes constant time -> O(1).
        if (!row[id_index]) continue;
        // Retrieving the key value takes constant time -> O(1).
        const key: string = row[id_index];
        // Initializing an empty object takes constant time -> O(1).
        data[key] = { cells: [] }

        // Iterating through each cell in the row takes linear time -> O(n), where n is the number of cells in the row.
        for (const [cell_index, cell] of Array.from(row.entries())) {
            // Adding the cell to the data object takes constant time -> O(1).
            data[key].cells.push({ column: columns[cell_index], value: Serialize.execute(cell) });
        }
    }

    return data;
}
```

### **Parsing Excel Data $O(n^2)$:**
Similarly to parsing CSV data, the time complexity of parsing Excel data is dominated by the nested loop that iterates through each row and each cell, resulting in a time complexity of $O(n^2)$, where $n$ is the total number of cells in the Excel data.

#### Time Complexity Analysis
* Retrieving all the values from the active worksheet's used range takes linear time, $O(n)$, where n is the number of cells in the used range.
* Converting each cell value to a string, trimming it, and converting it to a string again takes constant time, $O(1)$.
* Initializing an empty object takes constant time, $O(1)$.
* Finding the index of the key in the columns array takes linear time, $O(n)$.
* Iterating over each row takes linear time, $O(n)$, where n is the number of rows.
* Initializing a new key in the data object takes constant time, $O(1)$.
* Initializing a new object in the data object takes constant time, $O(1)$.
* Iterating over each cell in the row takes linear time, $O(n)$, where n is the number of cells in the row.
* Pushing a new cell object to the data[key].cells array takes constant time, $O(1)$.

```typescript
function parseExcel(workbook: ExcelScript.Workbook, key: string): { columns: Set<string>; rows: ParsedCollection; } {
    // Retrieving all the values from the active worksheet's used range takes linear time -> O(n), where n is the number of cells in the used range.
    const [columns, ...rows]: string[][] = workbook.getActiveWorksheet().getUsedRange().getValues().map(
        // Converting each cell value to a string, trimming it, and converting it to a string again takes constant time -> O(1).
        (row: (string | number | boolean)[]) => row.map((cell) => `${cell ?? ''}`.trim().toString())
    );

    // Initializing an empty object takes constant time -> O(1).
    const data: ParsedCollection = {};
    // Finding the index of the key in the columns array takes linear time -> O(n).
    const id_index: number = columns.indexOf(key);

    // Iterating over each row takes linear time -> O(n), where n is the number of rows.
    for (const [row_index, row] of Array.from(rows.entries())) {
        // Initializing a new key in the data object takes constant time -> O(1).
        const key: string = row[id_index];
        // Initializing a new object in the data object takes constant time -> O(1).
        data[key] = { row: row_index + 1, cells: [] }
        // Iterating over each cell in the row takes linear time -> O(n), where n is the number of cells in the row.
        for (const [cell_index, cell] of Array.from(row.entries())) {
            // Pushing a new cell object to the data[key].cells array takes constant time -> O(1).
            data[key].cells.push({ column: columns[cell_index], value: cell });
        }
    }

    return { columns: new Set(columns), rows: data };
}
```

### **Partitioning Data $O(n^2)$:**
The partition function takes CSV, Excel, and Excel columns as input parameters. It performs a partitioning operation to identify rows that need to be added or updated in the Excel worksheet based on the differences between the CSV and Excel data. The time complexity of the partition function is dominated by the nested loop that iterates through each key in the CSV data and checks if the key exists in the Excel data, resulting in a time complexity of $O(n^2)$, where $n$ is the number of keys in the CSV data.

#### Time Complexity Analysis
* Initializing an empty array takes constant time, $O(1)$.
* Iterating over each key in the CSV data takes linear time, $O(n)$, where n is the number of keys in the CSV data.
* Checking if the key exists in the Excel data takes constant time, $O(1)$.
* Retrieving the row index from the Excel data takes constant time, $O(1)$.
* Initializing an empty object takes constant time, $O(1)$.
* Iterating over each cell in the row takes linear time, $O(n)$, where n is the number of cells in the row.
* Pushing a new cell object to the row.cells array takes constant time, $O(1)$.
* Pushing a new row object to the rowsToAdd array takes constant time, $O(1)$.
* Pushing a new row object to the rowsToUpdate array takes constant time, $O(1)$.

```typescript
    function partition(
        csvData: Parser.ParsedCollection,
        excelData: Parser.ParsedCollection,
        excelColumns: Map<string, number>
    ): { rowsToAdd: SheetRow[]; rowsToUpdate: SheetRow[] } {
        // Initializing an empty array takes constant time -> O(1).
        const rowsToAdd: SheetRow[] = [];
        // Initializing an empty array takes constant time -> O(1).
        const rowsToUpdate: SheetRow[] = [];

        // Iterating over each key in the CSV data takes linear time -> O(n), where n is the number of keys in the CSV data.
        for (const [key, value] of Object.entries(csvData)) {
            // Checking if the key exists in the Excel data takes constant time -> O(1).
            if (excelData[key]) {
                // Retrieving the row index from the Excel data takes constant time -> O(1).
                const row_index: number = excelData[key].row ?? 0;
                // Initializing an empty array takes constant time -> O(1).
                const cells: SheetRow["cells"] = [];

                // Iterating over each cell in the CSV data takes linear time -> O(n), where n is the number of cells in the CSV data.
                for (const cell of value.cells) {
                    // Retrieving the cell from the Excel data takes constant time -> O(n), where n is the number of cells in the Excel data.
                    const excel_cell: Parser.ParsedData | undefined = excelData[key].cells.find((excel_cell) => excel_cell.column === cell.column);

                    // Checking if the cell exists in the Excel data takes constant time -> O(1).
                    if (excel_cell && excel_cell.value !== cell.value) {
                        // Retrieving the cell index from the Excel columns map takes constant time -> O(1).
                        const cell_index: number | undefined = excelColumns.get(cell.column);
                        // Checking if the cell index exists in the Excel columns map takes constant time -> O(1).
                        if (cell_index !== undefined) {
                            // Pushing a new cell object to the cells array takes constant time -> O(1).
                            cells.push({ cell_index, value: cell.value });
                        }
                    }
                }

                // Pushing a new row object to the rowsToUpdate array takes constant time -> O(1).
                rowsToUpdate.push({ row_index, cells });
            } else {
                // Retrieving the row index from the Excel data takes constant time -> O(1).
                const row_index: number = Object.keys(excelData).length + rowsToAdd.length + 1;
                // Initializing an empty array takes constant time -> O(1).
                const cells: SheetRow["cells"] = [];

                // Iterating over each cell in the CSV data takes linear time -> O(n), where n is the number of cells in the CSV data.
                for (const cell of value.cells) {
                    // Retrieving the cell index from the Excel columns map takes constant time -> O(1).
                    const cell_index: number | undefined = excelColumns.get(cell.column);
                    // Checking if the cell index exists in the Excel columns map takes constant time -> O(1).
                    if (cell_index !== undefined) {
                        // Pushing a new cell object to the cells array takes constant time -> O(1).
                        cells.push({ cell_index, value: cell.value });
                    }
                }

                // Pushing a new row object to the rowsToAdd array takes constant time -> O(1).
                rowsToAdd.push({ row_index, cells });
            }
        }

        return { rowsToAdd, rowsToUpdate };
    }
```