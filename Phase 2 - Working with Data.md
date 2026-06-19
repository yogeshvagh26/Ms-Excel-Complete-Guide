# Phase 2: Working with Data 

Welcome to Phase 2! Now that you have a solid foundation in Excel's interface and basic operations, we'll dive into one of the most crucial skills: manipulating and managing data. Real-world data is rarely perfect or organized exactly how you need it. This phase teaches you the tools to clean, organize, extract, and control your data effectively.

Take your time, follow the demonstrations, and complete the exercises. Each concept builds on the previous one.

---

## Topic 1: Data Types

### Concept Explanation

#### What are Data Types?

In Excel, every piece of data you enter is stored as a specific "type." Excel uses these types to determine what you can do with the data—whether you can calculate it, sort it chronologically, or just display it as text.

The four primary data types are:

1. **Text (String)**: Any combination of letters, numbers, and symbols that Excel treats as plain text.

    * **Examples**: "John Doe", "Product A123", "New York"

    * **Default Alignment**: Left-aligned

    * **Behavior**: You cannot perform math on text values.

    ---


2. **Numbers (Numeric)**: Digits that can be used in calculations.

    * **Examples**: 1500, 0.25, -50, 3.14159

    * **Default Alignment**: Right-aligned

    * **Behavior**: You can add, subtract, average, etc. (Note: A number with a leading zero, like "0123", is usually treated as text unless formatted specially, because the zero would disappear).

    ---

3. **Dates and Times**: Special numeric values representing a specific point in time. Excel stores dates as serial numbers (e.g., Jan 1, 1900 = 1).

    * **Examples**: 15-Jan-2024, 3:30 PM

    * **Default Alignment**: Right-aligned

    * **Behavior**: You can calculate the difference between dates, add days, etc.

    ---

4. **Boolean (Logical)**: Represents TRUE or FALSE values, typically returned by logical formulas.

    * **Examples**: TRUE, FALSE

    * **Behavior**: Used in conditions and decision-making formulas.

---

### How to Identify Data Types:

* Look at the alignment (Left=Text, Right=Numbers/Dates).

* Look at the Number Format dropdown in the Home tab.

* Look for the green triangle in the top-left corner of a cell—this indicates a potential data type mismatch (e.g., a number stored as text).
---

### Importance and Real-World Use Cases

#### Why Data Types Matter:


* **Accurate Calculations**: If a "number" is stored as text, SUM will ignore it, leading to incorrect totals.

* **Correct Sorting**: If dates are stored as text, they will sort alphabetically (e.g., April, August, February) rather than chronologically (Jan, Feb, Mar).

* **Proper Filtering**: Filters work differently for text vs. numbers.

* **Database Integrity**: Ensuring the correct type prevents errors in reports.

#### Real-World Scenarios:

* **Financial Analyst**: If imported bank statements have numbers formatted as text, the entire budget report will be wrong.

* **HR Manager**: Sorting employee start dates as dates ensures seniority reports are correct.

* **Sales Team**: Filtering by sales amount requires numeric data types.

---

### Step-by-Step Demonstration

#### 1. Identifying Data Types via Alignment:

1. Enter "Hello" in A1 (Left-aligned → Text).

2. Enter "100" in B1 (Right-aligned → Number).

3. Enter "1/1/2024" in C1 (Right-aligned → Date).


#### 2. Converting Text to Numbers (The Green Triangle Method):

1. Type '00123 (apostrophe forces text) in A1.

2. Notice the green triangle.

3. Click the yellow diamond that appears.

4. Select "Convert to Number".


#### 3. Changing Data Types via Formatting (Does NOT change the value, just how it looks):

1. Right-click a cell → Format Cells.

2. Go to the Number tab.

3. Select "Number", "Currency", "Date", or "Text".

4. Warning: Formatting a cell as "Text" after typing a number will not convert it back. You must re-enter it.

---

### Syntax or Rules

#### Rules for Entering Data:

* **Text**: Can contain any characters. To force text (e.g., for ID numbers starting with 0), type an apostrophe `'` before the number, e.g., `'00123`.

* **Numbers**: Only digits, decimals, and negative signs. No commas (Excel treats them as thousands separators, which is okay).

* **Dates**: Use slashes (/) or hyphens (-). Excel recognizes `1/15/2024` or `15-Jan-2024`.

* **Times**: Use a colon (e.g., `3:30 PM`).

---

### Multiple Practical Examples

#### Example 1: Fixing Imported Data

You import a CSV file and all numbers are left-aligned.

* **Solution**: Select the column, go to Data → Text to Columns → Finish (this forces Excel to re-evaluate the type).

#### Example 2: Creating an Employee ID

You need IDs like "E001" and "E002".

* **Solution**: Type `'E001` (the apostrophe tells Excel it's text).

#### Example 3: Calculating Age from Date of Birth

If DOBs are stored as text, `=YEAR(TODAY())-YEAR(A1)` will fail.

---

### Practice Exercises

#### Exercise 1: Identifying Types
Enter the following data and identify (by alignment) what type each is:

* "123 Main St"

* 12345

* 12/25/2024

* 45.50(withthe sign)

#### Exercise 2: Converting Text to Numbers


1. In cell A1, type '1000.

2. In cell A2, type 2000.

3. In cell A3, write =SUM(A1:A2). Note the result (it only sums 2000).

4. Convert A1 to a number. Note the sum now becomes 3000.

---

## Topic 2: Copy, Paste, and Fill Options


### Concept Explanation

#### The Power of Copy, Paste, and Fill

Copying and pasting in Excel is far more powerful than in Word. You don't just paste text; you paste formulas, values, formats, and even transpose data (switching rows to columns).

#### Key Concepts:

* **Copy (Ctrl+C)**: Duplicates the source cell(s).

* **Paste (Ctrl+V)**: The standard paste. In Excel, this pastes everything (value, formula, formatting, comments).

* **Paste Special**: The magic tool. Allows you to choose what to paste.

* **Fill Handle**: The small black square at the bottom-right of a selected cell. Drag it down/right to copy the cell or fill a series.

#### Paste Special Options (Most Common):

* **Values (V)**: Pastes only the resulting value (no formula, no formatting). Crucial for copying calculated results without copying the logic.

* **Formats (T)**: Pastes only the formatting (like a paintbrush).

* **Transpose (E)**: Switches rows to columns and vice versa.

* **Multiply/Add (Operation)**: Allows you to paste values and multiply or add them to existing numbers simultaneously.

#### Fill Handle Options:

* **Copy Cells**: Repeats the exact same value.

* **Fill Series**: Increments numbers/dates (e.g., 1, 2, 3 or Jan, Feb, Mar).

* **Fill Formatting Only**: Copies only the cell color/font.

* **Flash Fill (Covered later)**: Fills based on a pattern.

---

### Importance and Real-World Use Cases


* **Reporting**: You calculate totals with formulas, but need to send only the final numbers to a client. You use Paste Values to remove the formulas.

* **Data Restructuring**: You have data in rows but need it in columns for a chart. You use Transpose.

* **Budget Updates**: You need to increase all prices by 10%. You type 1.1, copy it, select the price range, and use Paste Special → Multiply.

* **Templates**: You copy formatting from a perfect report to a new one using Paste Formats.

---

### Step-by-Step Demonstration

#### 1. Paste Values (Crucial Skill!)


1. In A1, type `10`. In B1, type `20`. In C1, write `=A1+B1` (result is `30`).

2. Right-click C1 → Copy (or Ctrl+C).

3. Right-click D1 → Paste Options → Values (the 123 icon).

4. Result: D1 shows `30`, but the formula bar is empty. It's just a number now.

#### 2. Transpose (Rotate Data)

1. Type "January" in A1, "February" in B1, "March" in C1.

2. Select A1:C1 → Copy.

3. Right-click A3 → Paste Special → Check "Transpose" → OK.

4. Result: January, February, March now appear vertically in rows A3, A4, A5.


#### 3. Using the Fill Handle for Series

1. Type "Monday" in A1.

2. Click the cell, hover over the bottom-right corner until the cursor turns into a thin black plus (+).

3. Drag down to A7. Result: Tuesday, Wednesday, Thursday... Sunday.

#### 4. Incrementing Numbers

1. Type `1` in A1. Hold Ctrl + drag down. Result: 1, 2, 3, 4...

2. _Tip_: If you just drag without Ctrl, it will copy "1" repeatedly.

---

### Syntax or Rules


#### Fill Handle Rules:

* **Text without number**: Drag copies the text (e.g., "Apple" stays "Apple").

* **Text with number**: Drag increments the number (e.g., "Product 1" → "Product 2").

* **Dates**: Drag increments by day.

* **Custom Lists**: Excel recognizes months and weekdays by default.

---

### Practice Exercises

#### Exercise 1: Paste Values

1. Create a formula in A1: `=10*5`.

2. Copy the result and paste it as a value into B1.

3. Delete A1. Notice B1 stays `50`.

#### Exercise 2: Transpose

1. Create a list of 5 product names in a row (A1:E1).

2. Transpose them to a column starting at A5.

#### Exercise 3: Fill Series

1. Create a list of even numbers from 2 to 20 (Hint: type 2 and 4, select both, then drag).

2. Create a list of dates for the next 30 days starting from today.

---

## Topic 3: Sorting Data

### Concept Explanation

#### What is Sorting?

Sorting is arranging your data in a specific order—alphabetically (A to Z, Z to A), numerically (smallest to largest, largest to smallest), or by date (oldest to newest, newest to oldest).

#### Levels of Sorting:

* **Single Level**: Sort by one column (e.g., sort customers by Last Name).

* **Multiple Levels (Custom Sort)**: Sort by primary column, then secondary column (e.g., sort by Department, and within each department, sort by Salary).

---

### Importance and Real-World Use Cases

* **Sales Analysis**: Sorting products by highest sales to identify top performers.

* **HR Management**: Sorting employees by hire date to find seniority.

* **Inventory**: Sorting stock by quantity (lowest first) to identify items needing reorder.

* **Daily Tasks**: Sorting a task list by due date to prioritize.

---

### Step-by-Step Demonstration

#### 1. Simple Sort (A-Z / Smallest to Largest)

1. Select any cell within your data range (e.g., a column of names).

2. Go to the Data tab.

3. Click AZ (Sort Ascending) or ZA (Sort Descending).

4. Excel Warning: If your data has a header row, Excel will usually detect it and ask if you want to expand the selection. Always expand to keep rows intact.

#### 2. Custom Sort (Multiple Levels)

1. Go to Data → Sort.

2. In the Sort dialog box:

    * **Column**: Choose "Department".

    * **Sort On**: Values.

    * **Order**: A to Z.

3. Click Add Level.

4. Then by: Choose "Salary".

    * **Order**: Largest to Smallest.

5. Click OK.

    * **Result**: The data is grouped by Department, and within each department, employees are listed from highest salary to lowest.

#### 3. Sort by Cell Color (Advanced)

1. Data → Sort.

2. Under "Sort On", select "Cell Color".

3. Choose the color you want to bring to the top (e.g., red highlighted cells for urgent issues).

---

### Syntax or Rules

#### Critical Sorting Rule:

* **ALWAYS include the header row in the selection, but tell Excel "My data has headers"**.

* If you select only one column and sort, the rest of the row will not move, and you will **mix up your data**. Always select the entire table or let Excel expand the selection.

---

### Practice Exercises

#### Exercise 1: Sorting Practice
Create this table:

| Name | City | Sales |
|------|------|-------|
| John | NY   | 500   |
| Anna | LA   | 700   |
| Bob  | NY   | 400   |
| Mary | LA   | 900   |


1. Sort the data by City (A-Z).

2. Sort the data by Sales (Largest to Smallest).

3. Sort by City (A-Z), and then by Sales (Largest to Smallest).

---

## Topic 4: Filtering Data

### Concept Explanation

#### What is Filtering?

Filtering is temporarily hiding rows that do not meet specific criteria, allowing you to focus on a subset of your data. Unlike sorting, which rearranges data, filtering just hides the rest.

#### How it works:

* When you enable filters (Data → Filter), dropdown arrows appear in the header row.

* You can select specific items (e.g., only "NY"), or use Number/Text Filters (e.g., "Greater than 500", "Begins with A").

#### Advanced Filters:

* **Text Filters**: Contains, Does Not Contain, Begins With, Ends With.

* **Number Filters**: Greater Than, Less Than, Between, Top 10.

* **Date Filters**: This Week, Last Month, Quarter, Year-to-Date.

* **Color Filters**: Filter by cell color or font color.

---

### Importance and Real-World Use Cases

* **Sales Manager**: Filter to see only sales made in the "West" region.

* **Accountant**: Filter transactions greater than $10,000 to identify large expenses.

* **Marketing**: Filter customers from the "Current Quarter" to analyze recent trends.

* **HR**: Filter employees by department to send department-specific emails.

---

### Step-by-Step Demonstration

#### 1. Applying a Simple Filter

1. Click anywhere inside your data table.

2. Go to Data → Filter (or press Ctrl+Shift+L).

3. Click the dropdown arrow in the "City" column.

4. Uncheck "Select All".

5. Check only "NY".

6. Excel now shows only rows where City is NY.


#### 2. Using Number Filters

1. Click the dropdown arrow in the "Sales" column.


2. Select Number Filters → Greater Than.

3. Enter `600`.

4. Click OK. Only sales > 600 appear.

#### 3. Clearing Filters

1. Click the Clear button on the Data tab (next to Filter) to show all data.


2. Or click the dropdown arrow and select "Clear Filter From [Column]".


#### 4. Filter by Selection (Right-click method)

1. Right-click a cell containing "LA".

2. Choose Filter → Filter by Selected Cell's Value. This instantly filters the entire list to "LA".

---

### Syntax or Rules

#### Filtering Rules:

* Filters work only on contiguous ranges (no blank rows in the middle of your data).

* Hidden rows are not deleted—they just aren't visible.

* Subtotals and totals at the bottom of a table can be excluded from filtering by leaving a blank row before the totals.

* To apply filters to multiple columns simultaneously, just apply the filter to each.

---


### Practice Exercises

#### Exercise 1: Filter Practice
Use the same table from Sorting.

1. Enable filters.

2. Filter to show only "LA".

3. Filter to show only Sales greater than 400.

4. Filter to show Sales between 300 and 800.

5. Clear all filters.


### Exercise 2: Text Filter
Create a list of customer emails:

* john@abc.com, mary@xyz.com, bob@abc.com

1. Filter to show only emails ending with "abc.com" (Text Filter → Ends With).

---

## Topic 5: Find and Replace

### Concept Explanation

#### What is Find and Replace?
This is the search engine within your Excel workbook. You can find specific data (text, numbers, formulas) and optionally replace it with something else.

#### Why it's powerful:

* Find specific customer names in a huge database.

* Replace "2023" with "2024" everywhere in the workbook.

* Find and correct spelling mistakes.

* Go to specific cells or ranges quickly.

* Find cells with specific formatting (e.g., find all yellow cells).

---

### Importance and Real-World Use Cases

* **Data Cleanup**: Replacing old department names with new ones (e.g., "HR" → "Human Resources").

* **Updating Reports**: Changing the year from 2023 to 2024 in hundreds of cells.

* **Fixing Errors**: Finding and correcting typos (e.g., "Manger" → "Manager").

* **Navigation**: Jumping directly to a specific cell (e.g., Go To Z1000).

---
### Step-by-Step Demonstration

#### 1. Basic Find

1. Press Ctrl+F.

2. Type "Bob" in the Find what box.

3. Click "Find Next" to cycle through each occurrence.

#### 2. Basic Replace

1. Press Ctrl+H (Find & Replace shortcut).

2. Type "Sales" in Find what.

3. Type "Marketing" in Replace with.

4. Click "Replace All" to change everything, or "Replace" to do one at a time.

#### 3. Advanced Options (Important!)

1. In Find/Replace, click Options >>.

2. Within: Choose Sheet or Workbook (search all tabs).

3. Search: By Rows or Columns.

4. Look in: Formulas, Values, or Comments.

5. Match case: "apple" ≠ "Apple".

6. Match entire cell contents: Prevents changing "123" inside "4123".


#### 4. Finding Formatting

1. In Find/Replace, click the dropdown next to "Format" (if no format, click the arrow).

2. Choose "Choose Format From Cell".

3. Click a cell with a yellow fill.

4. Click "Find All" to see every cell with a yellow fill.

---

### Practice Exercises

#### Exercise 1: Replace


1. Type "2023" in cells A1, B2, C3.

2. Use Replace to change all "2023" to "2024".

3. Ensure the change happened correctly.


#### Exercise 2: Advanced Find

1. Enter "cat" and "Cat" in different cells.

2. Use Find with "Match case" enabled. Find only "cat".

3. Now use "Match entire cell contents" to find only cells that contain exactly "cat" (not "catalog").


---
 
## Topic 6: Text to Columns

### Concept Explanation

#### What is Text to Columns?

This tool takes a single column of data and splits it into multiple columns based on a delimiter (a character that separates values) or a fixed width.

#### Two Methods:


1. **Delimited**: Splits data based on characters like commas, tabs, spaces, or semicolons.


    * Example: "Doe, John" → Column A: "Doe", Column B: "John".

2. **Fixed Width**: Splits data by adding vertical lines at specific character positions.


    * Example: "00123ABC" → Column A: "001", Column B: "23", Column C: "ABC" (if you set fixed widths).

---

### Importance and Real-World Use Cases





