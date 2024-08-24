**Project Overview:**

This VBA project consists of three key macros designed to streamline and automate data processing tasks for the PerfectMind system and related sales reports. 

### Macros Included:

1. **CleanPerfectMind:**
   - **Purpose:** Cleans and restructures data in the "PMQT" worksheet. It retains only specific columns, removes unwanted entries like "--None--," concatenates and renames columns, and rearranges data for better usability.
   - **Steps Performed:**
     1. Retains the columns "Name," "Attribute 1," and "Available Quantity."
     2. Removes occurrences of "--None--."
     3. Concatenates columns A and B into a new column A, then deletes the original columns.
     4. Renames the new columns and switches their positions.

2. **CleanAndSummarizeSales:**
   - **Purpose:** Cleans and summarizes sales data in the "Sales" worksheet. It rearranges columns, concatenates data, and provides a summary of total sales per item.
   - **Steps Performed:**
     1. Retains the columns "ProductName," "Attribute1," and "Sold."
     2. Switches columns to align data for concatenation.
     3. Concatenates columns A and B into a new column A, deletes the originals, and summarizes sales data into columns C and D.

3. **Controlreport:**
   - **Purpose:** Matches data across the "Control Report," "Sales," and "PMQT" worksheets. It calculates inventory differences by pulling and processing data from these sources.
   - **Steps Performed:**
     1. Matches data from the "Control Report" with the "Sales" and "PMQT" sheets.
     2. Transfers relevant data between sheets.
     3. Calculates the difference between physical counts, PerfectMind quantities, and sales, placing the result in the "Control Report" worksheet.

---

**Important Note:**
After running the `CleanPerfectMind` and `CleanAndSummarizeSales` macros, **make sure to run the `Controlreport` macro last**. This ensures that all necessary data cleaning, restructuring, and summarizing are completed before performing the final data matching and inventory difference calculations.

