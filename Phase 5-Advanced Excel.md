# Phase 5: Advanced Excel

Welcome to Phase 5! This is the final frontier of core Excel mastery. While basic and intermediate Excel makes you productive, Advanced Excel makes you indispensable. This phase is about building robust, dynamic, and error-proof systems. We will cover formulas that adapt automatically, powerful dynamic arrays that "spill" results, advanced lookup tricks to find anything, and professional error handling.

_Disclaimer: Some Dynamic Array functions (`FILTER`, `UNIQUE`, `SORT`) are only available in Excel 365 and Excel 2021. If you have an older version, I will provide alternative methods._

---

## Topic 1: Advanced Formulas

### Concept Explanation

#### What are Advanced Formulas?

An advanced formula isn't necessarily a single exotic function; rather, it is a combination of functions used to solve complex, multi-step problems in a single cell. It involves using arithmetic operations with logical tests, handling text and numbers simultaneously, and creating flexible ranges.

#### Key Patterns:

* **Boolean Logic**: Using multiplication (`*`) and addition (`+`) inside formulas to mimic `AND` and `OR` without writing them.

* **Example**: `=(A1>10)*(B1="Yes")` returns 1 if both are true, else 0.

* **Array Operations**: Performing calculations on multiple values at once before summarizing.

* **Dynamic Range Detection**: Using functions like `COUNTA` to automatically adjust ranges so formulas work as data grows.

### Importance and Real-World Use Cases

* **Dynamic Reporting**: Creating a summary that updates automatically when new data is appended (e.g., a rolling 12-month sales summary).

* **Weighted Averages**: Calculating average cost of inventory where different batches have different prices.

* **Complex Incentive Calculations**: Calculating bonuses that depend on multiple tiers and department KPIs.

### Step-by-Step Demonstration

#### 1. Weighted Average Calculation (Without Helper Columns)

**Scenario:** 

> You have Units (A2:A10) and Price (B2:B10). You need the average price paid, weighted by units.

**Formula:** `=SUMPRODUCT(A2:A10, B2:B10) / SUM(A2:A10)`

* `SUMPRODUCT` multiplies Units by Price, sums the results, and divides by total units.

#### 2. Running Total (Cumulative Sum) with Dynamic Reference

**Scenario:** 

> List of daily sales in B2:B100. You want the running total in C2.

**Formula** in **C2**: `=SUM($B$2:B2)`

* Lock the start ($B$2), let the end (B2) be relative. Drag down. It sums from the top to the current row.

#### 3. Summing Every Nth Row

**Scenario**: Budget data where line items alternate with subtotals, and you want to sum only the values in every 3rd row.

**Formula**: `=SUMPRODUCT((MOD(ROW(A1:A100),3)=0)*A1:A100)`

### Practice Exercises

**Exercise 1: Weighted Average**

* You bought 100 shares at $10, 200 shares at $15, and 150 shares at $12. Calculate the average price you paid per share.

**Exercise 2: Dynamic Range Last 5 Entries**

* Assuming Column A has numeric data extending downward, write a formula to average the last 5 entries (hint: combine `AVERAGE`, `OFFSET`, and `COUNTA`).

---

## Topic 2: Nested Functions

### Concept Explanation

#### What are Nested Functions?

Nesting means putting one function inside another function as an argument. This creates a "formula sandwich." For example,` =IF(ISBLANK(A1), "Empty", A1)` uses `ISBLANK` inside `IF`.

The "75 Levels" Rule: Modern Excel allows up to 64 levels of nested functions, but if you are writing more than 3-4 levels, it becomes unreadable and difficult to debug. Best practice is to split into helper columns or use `LET` (Excel 365) to define variables.

#### Common Nesting Combinations:

* `IF(OR(...), ...)`

* `IF(AND(...), ...)`

* `VLOOKUP(LEFT(...), ...)`

* `SUM(IF(...)) (Array formula without Ctrl+Shift+Enter in modern Excel).`

----

### Importance and Real-World Use Cases

* **Data Classification**: Assigning a category based on multiple text conditions (e.g., `=IF(OR(A1="NY", A1="CA"), "East/West", "Central")`).

* **Complex Validation**: Checking if a date is a weekday before calculating a deadline.

* **Data Extraction**: Extracting a number from within a text string using `MID`, `FIND`, and `ISNUMBER`.

---

### Step-by-Step Demonstration

#### 1. Nested IF with AND (Employee Performance)
* **Scenario**: If Sales > 100,000 AND Complaints < 5, Bonus = "High"; else "Low".

* **Formula**: `=IF(AND(B2>100000, C2<5), "High", "Low")`


#### 2. VLOOKUP with LEFT (Extract and Lookup)

* **Scenario**: You have full Product Codes like "P123-ABC" in A1. The lookup table only has the first 4 characters ("P123").

* **Formula**: `=VLOOKUP(LEFT(A1, 4), LookupTable!A:B, 2, FALSE)`


#### 3. The LET Function (Game Changer for Readability - Excel 365)

Instead of repeating a calculation, define it.

**Formula**: `=LET(x, A1*B1, y, C1/D1, x + y)`


> Why: It avoids calculating the same thing multiple times and makes formulas readable.

---

### Practice Exercises

**Exercise 1: Nested IF**

* Create a grade calculator. If Score >= 90 -> "A", >=80 -> "B", >=70 -> "C", else "D". Write a single nested IF formula.

**Exercise 2: ISERROR + VLOOKUP**

* Write a formula that looks up a value. If the lookup returns an error, show "Not Available". (We will cover this more formally in Error Handling).

----

## Topic 3: Dynamic Arrays (Excel 365/2021)

### Concept Explanation

#### What are Dynamic Arrays?

Before Dynamic Arrays, formulas could only return a value to **a single cell. Now, a single formula can return multiple values that "spill" down or across** into adjacent cells. This is the biggest shift in Excel in decades.

#### Core Dynamic Array Functions:

Core Dynamic Array Functions:

* `UNIQUE(range)`: Extracts all distinct values from a range/column.

* `SORT(range, sort_index, order)`: Sorts a range based on a column.

* `SORTBY(range, by_array, order)`: Sorts based on another array.

* `FILTER(range, include, [if_empty])`: Filters data based on criteria. This is revolutionary.

* `SEQUENCE(rows, cols, start, step)`: Generates a list of numbers.

* `RANDARRAY()`: Generates random numbers.

* `#SPILL Error`: This appears if something is blocking the cells where the formula wants to return results.

---

### Importance and Real-World Use Cases

* **Automated Lists**: Create a unique list of products that automatically updates when you add new products to the source.

* **Dynamic Reports**: Use FILTER to show only rows where "Status = Active" and SORT to order them by date—all without PivotTables or macros.

* **Data Validation Sources**: Use UNIQUE inside Data Validation to create self-updating dropdowns.


---

### Step-by-Step Demonstration

#### 1. Extracting Unique Values

1. Enter a list with duplicates in A1:A20.

2. In C1, type =UNIQUE(A1:A20) and press Enter.

3. Excel returns a list of unique values spilling down.

#### 2. Filtering Data with Criteria

**Scenario**: 

> Table with Headers in A1:C100 (Name, City, Sales).

1. In E1, type the City "NY".

2. In G1, type =FILTER(A1:C100, B1:B100=E1, "No data").

3. Press Enter. Excel spits out only the rows where City = NY.

#### 3. Combining FILTER and SORT

* `=SORT(FILTER(A1:C100, B1:B100="NY"), 3, -1)`

* This filters for NY and sorts the result by column 3 (Sales) in descending order (-1).

---

### Practice Exercises

**Exercise 1: Unique List**

* Create a list of 20 departments with duplicates. Use `UNIQUE` to generate a clean list in a new column.

**Exercise 2: Dynamic Filter**

* Using a dataset of Employees (Name, Department, Salary), use `FILTER` to display only the "Sales" department employees.
Challenge: Wrap it in `SORT` to sort the displayed list by Salary from Highest to Lowest.

---

## Topic 4: Advanced Lookup Techniques

### Concept Explanation

#### Going Beyond VLOOKUP

VLOOKUP is limited: it only looks to the right, requires a column index number, and breaks if you insert a column. XLOOKUP fixes this. But what about multiple conditions? We will conquer that.

#### Techniques:

1. **XLOOKUP with Concatenation**: Combine columns for multiple criteria.

2. **INDEX-MATCH-MATCH**: A classic two-way lookup (lookup a row AND a column).

3. **SUMPRODUCT for Lookups**: Non-array method to return a value based on multiple conditions.

---

### Importance and Real-World Use Cases

* **Inventory Systems**: Look up a price based on both "Product ID" and "Warehouse ID".

* **Timesheets**: Find an employee's shift based on "Date" and "Shift Code".

* **Matrix Tables**: Looking up an interest rate based on "Credit Score Range" (Rows) and "Loan Amount Range" (Columns).

---

### Step-by-Step Demonstration

#### 1. Multi-Condition Lookup (XLOOKUP or INDEX/MATCH)
**Scenario:** 

> Table A1:C10 has Product, Color, Price. You want Price for Product "Shirt" that is "Red".

**Using XLOOKUP:**

* `=XLOOKUP(1, (A2:A10="Shirt")*(B2:B10="Red"), C2:C10)`

* The product of logical tests creates an array of 1s and 0s. XLOOKUP looks for 1.

#### 2. Two-Way Lookup (Row & Column)
**Scenario:** 

> You have a table with Months in Row 1 (Jan, Feb) and Products in Column A (Apple, Banana). You want the Sales for Banana in February.

* **Formula:** =INDEX(B2:C3, MATCH("Banana", A2:A3, 0), MATCH("Feb", B1:C1, 0))

* INDEX takes the range. MATCH finds the row number. MATCH finds the column number.

#### 3. Wildcard Lookup (Partial Match)
**Scenario:** 

> You want to find a product containing "Pro" in the name.

* **Formula:** =XLOOKUP("*Pro*", A2:A10, B2:B10, "Not Found", 2) (The 2 indicates a wildcard match).

---

### Practice Exercises

**Exercise 1: Multi-Condition Lookup**

* You have a Price List: Product, Size (S/M/L), Price. Find the price for "T-Shirt" in size "L".

**Exercise 2: Two-Way Lookup**

* Create a multiplication table (or sales matrix). Use INDEX-MATCH-MATCH to retrieve the value at the intersection of "Row 5" and "Column 3".

---

## Topic 5: Named Ranges

### Concept Explanation

#### What are Named Ranges?

A Named Range is a descriptive name (e.g., "SalesData", "TaxRate") given to a cell, a range of cells, a formula, or a constant. Instead of typing `=SUM($B$2:$B$100)`, you type `=SUM(Sales)`.

**Scope:**

* **Workbook Scope**: The name is recognized across all sheets in the workbook.

* **Worksheet Scope**: The name is only recognized on that specific sheet.


#### Dynamic Named Ranges:

Using `OFFSET` or `INDEX`, you can make a Named Range that expands/shrinks automatically as your data changes. However, with the advent of Excel Tables, using Tables is now the preferred method for dynamic ranges.

---

### Importance and Real-World Use Cases

* **Readability**: `=Revenue - Costs` is infinitely easier to understand than `=A1 - B1`.

* **Navigation**: You can type "TaxRate" in the Name Box (top-left) and jump directly to it.

* **Data Validation**: Use Named Ranges as the source for dropdown lists (so they are dynamic).

* **Macros (VBA)**: Named Ranges make VBA code robust (won't break if rows are inserted).

---

### Step-by-Step Demonstration

#### 1. Creating a Named Range

1. Select cells B2:B20 (containing sales).

2. Click the Name Box (left of the Formula Bar).

3. Type `Sales` and press Enter.

4. Now write `=SUM(Sales)` anywhere. It works!

#### 2. Using the Name Manager (Ctrl+F3)

1. Press Ctrl+F3 to open the Name Manager.

2. Click New.

3. Name: `Discount`, Scope: Workbook.

4. Refers to: `=0.1` (Just a constant).

5. Now you can use `=Price * Discount` in formulas.

#### 3. Dynamic Named Range (Using OFFSET)

*Old school but useful: `=OFFSET(Sheet1!$A$1, 0, 0, COUNTA(Sheet1!$A:$A), 1)`*

* This refers to A1 down to the last non-empty cell in column A.

----

### Practice Exercises

**Exercise 1: Named Ranges for Tax**

1. Create a Named Range "Tax_Rate" with a value of 0.15.

2. Create a list of Invoices and calculate Tax using the Named Range.

**Exercise 2: Jumping**

1. Name cell Z100 as "Summary".

2. Now, type "Summary" into the Name Box and press Enter. Notice how you jump instantly.

---

## Topic 6: Error Handling

### Concept Explanation

#### What is Error Handling?

Data is messy. Formulas will often return errors like `#N/A` (Value not found) or `#DIV/0!` (Division by zero). Error handling functions catch these errors and replace them with something user-friendly, like a blank cell, "0", or "Not Found".

####  Core Functions:

* `IFERROR(value, value_if_error)`: Catches any error (#N/A, #VALUE!, #REF!, etc.).

* `IFNA(value, value_if_na)`: Catches only #N/A errors. Crucial for lookups!

* `ISERROR(value), ISNA(value)`: Return TRUE/FALSE if an error exists.

----

### Importance and Real-World Use Cases
* **Professional Reports**: No one wants to see #N/A in a board presentation. You want it to say "Not Found".

* **Dashboards**: Error-free dashboards look polished.

* **Preventing Chain Reactions**: If one formula breaks, it can break dependent formulas. Handling errors stops the spread.


---
### Step-by-Step Demonstration

#### 1. IFERROR for Division (Prevent #DIV/0!)
* **Scenario**: 

    >   A1 = 100, B1 = 0. =A1/B1 gives #DIV/0!.
Formula: =IFERROR(A1/B1, 0) → Returns 0 instead of an error.

#### 2. IFNA for VLOOKUP/XLOOKUP
* **Scenario**: 

    >   You are looking up a product that may not exist in the list.

* **Formula**: 

    `=IFNA(VLOOKUP(D1, A:B, 2, FALSE), "Product Not Found")`

* If it doesn't find it, it shows "Product Not Found". If there is a `#VALUE!` error (different type), `IFNA` will not catch it, allowing you to fix the structural error.

#### 3. Aggregating with Errors (SUM/AVERAGE)

* If a range contains errors, `SUM` returns an error.

* **Workaround**: `=SUM(IFERROR(A1:A10, 0))` (This is an array formula, or use `AGGREGATE`).

* **Pro Tip**: `=AGGREGATE(9, 6, A1:A10)` sums while ignoring errors (`9`=SUM, `6`=Ignore errors).

---

### Practice Exercises

**Exercise 1: Error-Proof VLOOKUP**

* Write a VLOOKUP that searches for an ID. If the ID is missing, display "Check ID".

**Exercise 2: Multiple Error Checks**

* Use IFERROR combined with IF to handle a scenario where a user enters text in a numeric field (causing #VALUE!), replacing it with "Input Error".

---

## Topic 7: Data Cleaning Techniques (Advanced)

### Concept Explanation

#### Beyond TRIM and PROPER

Advanced data cleaning involves correcting structural issues that basic functions can't fix. We use a combination of text functions, error handling, and logical tests to standardize messy data.

#### Advanced Techniques:

* **Removing Non-Printable Characters**: CLEAN removes characters that don't print (often imported from web/PDF).

* **Extracting Numbers from Text**: Combining MID, SEQUENCE, and ISNUMBER to pull digits out of an alphanumeric string.

* **Changing Case & Standardizing**: PROPER, UPPER, LOWER.

* **Substituting Multiple Values**: Using SUBSTITUTE nested to fix common typos ("NY" vs "N.Y.").

* **Separating First/Last Names with Multiple Spaces**: Using TRIM and FIND robustly.

### Importance and Real-World Use Cases

* **CRM Migration**: Before uploading a customer list to Salesforce, you must standardize phone numbers and addresses.

* **Financial Imports**: CSV files often import dates in text format (e.g., "20240115"). You need to convert them to real dates.

* **Data Science Prep**: Data must be clean before you put it into a PivotTable or Power BI.

---

### Step-by-Step Demonstration

**1. Extracting Numbers from a String**

**Scenario**: A

* 1 = "Order 12345 Status".

**Formula**: 

* `=TEXTJOIN("", TRUE, IF(ISNUMBER(--MID(A1, ROW($A$1:$A$100), 1)), MID(A1, ROW($A$1:$A$100), 1), ""))
`
* **Note**: This is a complex array formula (Dynamic Arrays handle it easier).

* **Modern Excel**: =REGEXEXTRACT is coming, but for now, we use TEXTJOIN with ISNUMBER.

**2. Converting Text Dates to Real Dates**

**Scenario**: 

* A1 contains "2024-01-15" (as text, left-aligned).

**Formula**: 

* `=DATEVALUE(A1)` → Converts to a real date (serial number). Format it as a date.


**3. Nested SUBSTITUTE (Cleaning Addresses)**

* **Scenario**: Addresses have "St" but you need "Street".

* **Formula**: =SUBSTITUTE(SUBSTITUTE(A1, "St", "Street"), "Ave", "Avenue")


**4. Using Flash Fill vs. Formulas**

_Reminder: Flash Fill is great for one-off cleanups. But formulas are dynamic. If your source data changes weekly, use formulas!_

---

### Practice Exercises
**Exercise 1: Clean Phone Numbers**

* A1 has " (123) 456-7890 ". Remove the spaces, parentheses, and dash to get just "1234567890". (Hint: Use SUBSTITUTE or CLEAN).

**Exercise 2: Extract Domain**

* Given emails in A1:A10 (e.g., "john@abc.com"), extract "abc.com" using a formula that finds the "@" position and uses RIGHT.


---

## Mini Quiz: Phase 5 - Advanced Excel

#### Section 1: Core Concepts

1. What is the primary benefit of using the `LET` function?
2. What does the `#SPILL!` error indicate?
3. What is the difference between `IFERROR` and `IFNA`?

#### Section 2: Advanced Lookups & Arrays

4. How do you perform a lookup with **two criteria** (e.g., Product AND Color) using `XLOOKUP`?
5. What two functions (used together) create a **Two-Way Lookup** (finding a value at the intersection of a specific row and column)?
6. Write a formula using `UNIQUE` and `SORT` together to list unique sales reps in alphabetical order.

#### Section 3: Named Ranges & Cleaning

7. What is the advantage of using a **Named Range** in a formula compared to a cell reference like `$A$1:$A$100`?
8. What function removes non-printable characters from imported data?
9. How do you prevent a division by zero error from showing `#DIV/0!`?

----

### Common Mistakes

1. **Hardcoding Values in Advanced Formulas:** Mistake: Writing `=A1*0.05` directly in a complex formula. Solution: Put `0.05` in a cell, name it "TaxRate", and reference the name. Update once, update everywhere.

2. **Over-Nesting (The "Spaghetti Formula"):** Mistake: Writing `=IF(IF(IF(...)))` with 15 levels. Solution: Break it down. Use helper columns or `LET`. Future you (and your colleagues) will thank you.

3. **Using VLOOKUP instead of XLOOKUP:** Mistake: Struggling with column indices and counting columns. Solution: Use `XLOOKUP` if available. It's simpler and more robust.

4. **Ignoring #SPILL Errors:** Mistake: Wondering why a dynamic array formula isn't working, only to find a stray value in the spill range. Solution: Clear the cells where the results are supposed to go.

5. **Using Merge Cells in Data:** Mistake: Merging cells across a Table or a range that will be used for lookup formulas. Merge kills sorting and PivotTables. Solution: Use "Center Across Selection" (*Format Cells* > *Alignment* > *Horizontal*) for visual centering without merging.

6. **Not Using Tables for Named Ranges:** Mistake: Creating dynamic `OFFSET` formulas for Named Ranges. Solution: Use Excel Tables. Table names (e.g., `Table1[Sales]`) auto-expand and are vastly easier.

### Best Practices

1. **Embrace LET:** In Excel 365, start every moderately complex formula with `LET` to define variables. It looks like code, and it's much easier to debug.

2. **Defensive Design:** Always wrap lookups in `IFNA`. Always protect divisions with `IFERROR`. Plan for messy data.

3. **Spill Range References:** Refer to the *first cell* of a dynamic array using the `#` operator. Example: `=SORT(...)` in A1. To reference the whole sorted list, use `A1#`.

4. **Data Validation + Named Ranges:** Combine `UNIQUE` (spilling) with a Named Range for dynamic dropdowns. In Data Validation, use `=UniqueList#` as the source.

5. **Consistent Naming:** Name ranges with underscores (e.g., `Sales_Data`) to avoid spaces. Be descriptive (`hr_Start_Date` vs `Date1`).

6. **Documentation:** If you write a super-advanced formula, add a comment using the `N()` function. Example: `=A1 + N("This adds the base salary")`.

---

## Interview Questions

### Intermediate to Advanced Level

**Q1: Explain the difference between `SUMIFS` and `SUMPRODUCT`. When would you use `SUMPRODUCT`?**

**A1:** `SUMIFS` is easier and faster for multiple criteria. `SUMPRODUCT` is more versatile because it can handle arrays, perform operations *within* the array (like multiplying two columns before summing), and works with array logic. Use `SUMPRODUCT` when you need to perform a calculation on the criteria (e.g., `(Sales*Quantity)`).

**Q2: What is a Dynamic Array and why is it important?**

**A2:** It's a formula that returns multiple values that automatically "spill" into adjacent cells. It's important because it eliminates the need for manual dragging, simplifies complex reports (like `FILTER`), and makes Excel feel more like a modern programming language.

**Q3: How do you perform a left-side lookup without XLOOKUP?**

**A3:** You use `INDEX` and `MATCH` together. `MATCH` finds the row number of the value in the column you are searching. `INDEX` then returns the value from a column to the left of that matched row.

**Q4: How can you avoid `#REF!` errors when deleting rows in a source range used by a formula?**

**A4:** Use Excel Tables with Structured References (`Table1[Sales]`). If you delete a row in a Table, the formula automatically adjusts the range reference. If you must use a range, use Dynamic Named Ranges.

**Q5: What is the `#` spill range operator and how do you use it?**

**A5:** If cell `A1` contains a spilling formula (e.g., `=UNIQUE(A:A)`), typing `A1#` refers to the entire spilled range. You can use `=SUM(A1#)` to sum a dynamic list that changes length.

**Q6: Can you use VLOOKUP with a dropdown list?**

**A6:** Yes. If you have a dropdown (Data Validation) in cell `D1` with product names, you can write `=VLOOKUP(D1, A:B, 2, FALSE)` to dynamically show the price of the selected product.

**Q7: How do you force Excel to treat a date stored as text as a real date?**

**A7:** Use the `DATEVALUE` function. If the text is `"2024-01-15"`, `=DATEVALUE(A1)` returns the serial number. Then format the cell as a date. Alternatively, use `VALUE` if it's a number stored as text.

**Q8: What is the `AGGREGATE` function and why is it better than `SUBTOTAL`?**

**A8:** `AGGREGATE` is a newer function that combines the functionality of `SUBTOTAL` but has options to ignore errors and hidden rows. For example, `=AGGREGATE(9, 6, A1:A10)` sums while ignoring any error values in the range.

---

## Assignment and Project Work: Data Consolidation and Reporting Engine

### Scenario: Enterprise Data Analyst

You have been tasked with creating a fully automated reporting engine. Your company receives a weekly raw sales export (messy, inconsistent, with missing data). Your job is to create a "Master Report" sheet that cleans, filters, and analyzes the data *without any manual intervention* (except pasting the new raw data).

**Dataset Setup (Create in a sheet named "Raw_Data"):**

Create a messy dataset with at least 50 rows. Include these columns and messiness:

* **Order ID:** Mixed case, e.g., "ord-101", "ORD-102"
* **Sales Rep:** Some with trailing spaces, e.g., " John ", "Mary "
* **Region:** Inconsistent, e.g., "East", "east", "EAST"
* **Product:** e.g., "Pro-X", "ProX", "Pro X"
* **Sales:** Numbers stored as text (e.g., '1000', '2000' - left aligned).
* **Profit:** Contains some errors (e.g., `#DIV/0!` in a few cells).
* **Date:** Stored as text like "2024-01-15".


**Tasks to Perform:**

1. **Data Cleaning (Create a "Cleaned_Data" sheet or use `LET` formulas):**
    * **Standardize Region:** Use `UPPER` and `TRIM` to make all Region entries exactly "EAST", "WEST".
    * **Clean Sales Rep:** Use `TRIM` to remove leading/trailing spaces.
    * **Fix Sales:** Convert the text numbers to actual numbers (use `VALUE` or multiply by 1).
    * **Fix Dates:** Convert text dates to real dates using `DATEVALUE`.
    * **Error Handling:** Where Profit has `#DIV/0!` or any error, replace it with `0` or a blank cell.
2. **Dynamic Array Summary (On a "Dashboard" sheet):**
    * Use `UNIQUE` to generate a dynamically updated list of all **Regions**.
    * Use `UNIQUE` to generate a dynamically updated list of all **Sales Reps** (without duplicates).
3. **Advanced Lookup & Filtering:**
    * Use `FILTER` to create a dynamic table showing only orders where **Sales > 5000 AND Region = "EAST"**.
    * Wrap this `FILTER` in a `SORT` to order these high-value east orders by Date descending.
4. **Named Ranges:**
    * Name the cleaned `Sales` column "Clean_Sales".
    * Name the cleaned `Profit` column "Clean_Profit".
    * Use these Named Ranges in a summary calculation at the top of the Dashboard: Total Sales, Average Profit.

5. **Nested Functions (Bonus - Excel 365):**
    * Create a single `LET` formula that outputs the total sales for the "West" region, but only if the sales rep is "Mary".

**Deliverables:**

* A workbook where simply pasting new data into "Raw_Data" automatically updates the "Cleaned_Data" and "Dashboard" outputs.
* All formulas must be fully dynamic (no static ranges).


---

## Summary and Revision Notes

### Phase 5: Key Concepts - Quick Revision

### 1. Advanced Formulas

* **Weighted Avg:** `SUMPRODUCT(A:A,B:B)/SUM(A:A)`.
* **Running Totals:** `SUM($B$2:B2)`.
* **LET:** Define variables (`=LET(x, A1+B1, x*2)`).

### 2. Nested Functions

* Nesting depth is limited to 64 levels. Use helper columns or `LET` for readability.
* Pattern: `=IF(AND(cond1, cond2), True, False)`.

### 3. Dynamic Arrays (Modern Excel)

* **Spawn:** `UNIQUE`, `SORT`, `SORTBY`, `FILTER`.
* **Operator:** Use `#` to reference the spill range (e.g., `A1#`).
* **Errror:** `#SPILL!` means blocked cells.

### 4. Advanced Lookups

* **Multi-Criteria XLOOKUP:** `=XLOOKUP(1, (Col1="X")*(Col2="Y"), Return_Range)`.
* **Two-Way Lookup:** `INDEX(Matrix, MATCH(RowLabel, RowHeaders,0), MATCH(ColLabel, ColHeaders,0))` .

### 5. Named Ranges

* **Create:** Name Box or Ctrl+F3.
* **Scopes:** Workbook (global) vs Worksheet (local).
* **Dynamic:** Use Tables instead of OFFSET.

### 6. Error Handling

* **IFERROR:** Catches everything (Use for calculations like division).
* **IFNA:** Catches `#N/A` only (Use for Lookups).
* **AGGREGATE:** Smart summing that can ignore errors.

### 7. Data Cleaning

* **TRIM, CLEAN, SUBSTITUTE.**
* **DATEVALUE / VALUE for type conversion.**
* **ISNUMBER for validating imports.**

---

**Congratulations!** You have completed the full core curriculum (Phases 1-5). You are now officially an **Advanced Excel User**.

_Keep building projects, and always look for ways to automate your daily tasks. Well done!_

---



<br/><br/><br/>
<center> <b>Happy Learning! 😊</b> </center>