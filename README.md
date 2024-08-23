Here’s a GitHub-style README for your VBA scripts:

---

# Excel VBA Macros

## Version 1.0

### Overview

This repository contains a set of VBA macros designed to automate various data processing tasks in Excel. The macros included help clean, summarize, and synchronize data across different worksheets, streamlining your workflow.

### Table of Contents

- [Macros Included](#macros-included)
  - [CleanSales](#cleansales)
  - [CleanPerfectMind](#cleanperfectmind)
  - [SummarizeSales](#summarizesales)
  - [ControlReport](#controlreport)
- [How to Use](#how-to-use)
- [Installation](#installation)
- [Error Handling](#error-handling)
- [Contributing](#contributing)
- [License](#license)

### Macros Included

#### CleanSales

Cleans and reformats the `Sales` worksheet by:
- Keeping only the specified columns: `ProductName`, `Attribute1`, and `Sold`.
- Switching the positions of columns B and C.
- Concatenating columns A and B into a new column.
- Renaming and switching columns for clarity.
- Replacing any occurrences of `--None--` with an empty string.

#### CleanPerfectMind

Cleans and reorganizes data in the `PerfectMind` worksheet by:
- Keeping only the specified columns: `Name`, `Attribute 1`, and `Available Quantity`.
- Removing occurrences of `--None--`.
- Concatenating columns A and B into a single column.
- Renaming columns and switching them for better readability.

#### SummarizeSales

Summarizes sales data by:
- Identifying unique items in the dataset.
- Calculating the total number of items sold.
- Displaying the summarized data in new columns C and D.

#### ControlReport

Synchronizes data between the `Control Report`, `Sales`, and `PMQT` worksheets by:
- Matching entries in `Control Report` column A with `Sales` column C, pulling corresponding data from `Sales` column D into `Control Report` column E.
- Matching entries in `Control Report` column A with `PMQT` column A, pulling corresponding data from `PMQT` column B into `Control Report` column D.

### How to Use

1. **Ensure the necessary worksheets are present:**
   - Your workbook should contain `Sales`, `PerfectMind`, `Control Report`, and `PMQT` sheets.

2. **Running a macro:**
   - Open the workbook in Excel.
   - Press `Alt + F8` to open the Macro dialog.
   - Select the desired macro and click `Run`.

3. **Backup your data:**
   - It’s advisable to create a backup of your data before running these macros, especially when working with large datasets.

### Installation

1. Download or clone this repository.
2. Open your Excel workbook.
3. Press `Alt + F11` to open the VBA editor.
4. Import the `.bas` files or copy the VBA code from the relevant script into a new module in your workbook.
5. Save your workbook as a macro-enabled workbook (`.xlsm`).

### Error Handling

The macros include basic error handling to manage cases where data does not match or is missing. If no match is found during data synchronization, the corresponding cell will display "No Match".

### Contributing

Contributions are welcome! If you have suggestions for improving these macros or want to add new features, please:
1. Fork the repository.
2. Create a new branch (`git checkout -b feature-branch`).
3. Make your changes.
4. Submit a pull request.

### License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

---

This GitHub-style README provides all necessary information for users to understand, install, and use the VBA macros, along with instructions for contributing and details about licensing.
