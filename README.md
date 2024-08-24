# VBA Scripts for Excel Automation

This repository contains VBA macros designed to automate various data processing tasks in Excel. The scripts are intended for use with specific worksheets and are tailored for tasks such as cleaning data, summarizing sales information, and calculating inventory differences.

## Scripts Overview

### 1. `CleanPerfectMind`

**Purpose**: This macro cleans and processes the "PMQT" worksheet. It retains only the specified columns, removes placeholder values, concatenates columns, and renames columns.

**Code**:
```vba
Sub CleanPerfectMind()
    ' ... [Code omitted for brevity] ...
End Sub
```

**Steps**:
1. Keep only specified columns: "Name", "Attribute 1", "Available Quantity".
2. Remove entries with value "--None--".
3. Concatenate columns A and B into column A, then delete original columns A and B.
4. Rename column B to "Item" and switch columns A and B.

**Usage**: Run this macro to clean and organize data in the "PMQT" worksheet according to the specified criteria.

### 2. `Controlreport`

**Purpose**: This macro processes data in the "Control Report", "Sales", and "PMQT" worksheets. It matches data between these sheets and calculates inventory differences.

**Code**:
```vba
Sub Controlreport()
    ' ... [Code omitted for brevity] ...
End Sub
```

**Steps**:
1. Match items in "Control Report" column A with "Sales" column C and pull corresponding data to column E.
2. Match items in "Control Report" column A with "PMQT" column A and pull corresponding data to column D.
3. Calculate the sum of Physical Count and PerfectMind Quantity, subtract Sales, and place the result in column F.

**Usage**: Execute this macro to update and analyze inventory data across the specified worksheets.

### 3. `CleanAndSummarizeSales`

**Purpose**: This macro processes the "Sales" worksheet by cleaning data, concatenating columns, and summarizing sales information.

**Code**:
```vba
Sub CleanAndSummarizeSales()
    ' ... [Code omitted for brevity] ...
End Sub
```

**Steps**:
1. Keep only specified columns: "ProductName", "Attribute1", "Sold".
2. Switch columns B and C, concatenate columns A and B into column A, and delete original columns A and B.
3. Replace occurrences of "--None--" with an empty string.
4. Summarize sales data and output unique items with total quantities sold.

**Usage**: Use this macro to clean and summarize sales data in the "Sales" worksheet.

## How to Use

1. Open your Excel workbook.
2. Press `ALT + F11` to open the VBA editor.
3. Insert a new module: `Insert > Module`.
4. Copy and paste the desired macro code into the module.
5. Press `F5` to run the macro or use `Tools > Macros` to select and execute the macro.

