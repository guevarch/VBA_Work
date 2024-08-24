# VBA Scripts for Excel Automation

This repository contains VBA macros developed to streamline and automate specific data processing tasks within Excel. These scripts are particularly useful for managing and analyzing data related to sales, inventory, and custom reports, helping to improve efficiency and accuracy in data handling.

## Overview

### `CleanPerfectMind`

**Purpose**:  
The `CleanPerfectMind` macro is designed to clean and reformat data within the "PMQT" worksheet, ensuring the data is organized and ready for further analysis or reporting. The macro performs a series of operations, including:

- **Column Filtering**: Retains only essential columns (e.g., "Name," "Attribute 1," "Available Quantity") while removing any unnecessary data.
- **Data Cleaning**: Removes placeholder values like "--None--" that may interfere with data interpretation.
- **Data Transformation**: Concatenates information from two columns into one, facilitating easier data analysis and ensuring that key details are grouped together.
- **Column Renaming and Rearranging**: Ensures columns are named appropriately and ordered in a way that aligns with the expected data structure.

**Usage**:  
This macro is useful when preparing the "PMQT" worksheet for inventory management or report generation, especially when dealing with large datasets that require consistent formatting and cleaning.

### `Controlreport`

**Purpose**:  
The `Controlreport` macro is intended to automate data matching and inventory calculations across multiple worksheets, including "Control Report," "Sales," and "PMQT." It enhances the accuracy and speed of complex data comparisons and calculations by:

- **Data Matching**: Automatically matches items between the "Control Report" and other relevant worksheets, pulling corresponding data to update the report.
- **Inventory Calculation**: Calculates the difference between physical counts, PerfectMind quantities, and sales figures, outputting the results for easy reference.
- **Error Handling**: Identifies and flags items that do not have corresponding matches in the other datasets, allowing for quick resolution of discrepancies.

**Usage**:  
Ideal for financial analysts, inventory managers, or anyone responsible for maintaining up-to-date and accurate inventory records. This macro simplifies the process of reconciling physical inventory counts with sales and system-generated quantities.

### `CleanAndSummarizeSales`

**Purpose**:  
The `CleanAndSummarizeSales` macro focuses on cleaning and summarizing sales data within the "Sales" worksheet. This macro ensures that sales data is not only clean but also summarized in a way that provides insights into total sales performance by:

- **Column Filtering and Cleaning**: Retains only the necessary columns (e.g., "ProductName," "Attribute1," "Sold") and cleanses the data by removing irrelevant entries.
- **Data Concatenation and Transformation**: Merges key details into a single column to create a more streamlined dataset.
- **Sales Summarization**: Identifies unique items and calculates the total quantity sold for each item, presenting the summarized data in a clear and organized manner.

**Usage**:  
This macro is particularly beneficial for sales managers, analysts, or anyone needing to prepare clean, summarized sales reports for presentations or further analysis. By automating data summarization, it reduces manual effort and enhances report accuracy.

## How to Implement

1. **Accessing the VBA Editor**:  
   - Open your Excel workbook.
   - Press `ALT + F11` to open the Visual Basic for Applications (VBA) editor.

2. **Adding a New Module**:  
   - In the VBA editor, insert a new module by navigating to `Insert > Module`.

3. **Pasting the Macro Code**:  
   - Copy the relevant macro code (from this repository) and paste it into the newly created module.

4. **Running the Macro**:  
   - To run the macro, press `F5` within the VBA editor or return to Excel, go to `Tools > Macros`, select the desired macro, and click `Run`.

## Best Practices

- **Backup Data**: Always ensure that your data is backed up before running any macros. This prevents data loss in case of unexpected issues.
- **Test on Sample Data**: Before applying macros to critical data, test them on a smaller dataset to confirm they work as expected.
- **Customize as Needed**: The macros are designed to be customizable. Adjust the code to better fit the specific needs of your data and reporting requirements.
