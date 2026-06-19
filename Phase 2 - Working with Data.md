# Phase 2: Working with Data 

Welcome to Phase 2! Now that you have a solid foundation in Excel's interface and basic operations, we'll dive into one of the most crucial skills: **manipulating and managing data**. Real-world data is rarely perfect or organized exactly how you need it. This phase teaches you the tools to clean, organize, extract, and control your data effectively.

Take your time, follow the demonstrations, and complete the exercises. Each concept builds on the previous one.

---

## Topic 1: Data Types

### Concept Explanation

#### What are Data Types?

In Excel, every piece of data you enter is stored as a specific "type." Excel uses these types to determine what you can do with the data—whether you can calculate it, sort it chronologically, or just display it as text.

The four primary data types are:

1. **Text (String)**: Any combination of letters, numbers, and symbols that Excel treats as plain text.

    * **Examples**: "John Doe", "Product A123", "New York"

    * **Default Alignment**: **Left-aligned**

    * **Behavior**: You cannot perform math on text values.

    ---


2. **Numbers (Numeric)**: Digits that can be used in calculations.

    * **Examples**: 1500, 0.25, -50, 3.14159

    * **Default Alignment**: **Right-aligned**

    * **Behavior**: You can add, subtract, average, etc. (Note: A number with a leading zero, like "0123", is usually treated as text unless formatted specially, because the zero would disappear).

    ---

3. **Dates and Times**: Special numeric values representing a specific point in time. Excel stores dates as serial numbers (e.g., Jan 1, 1900 = 1).

    * **Examples**: 15-Jan-2024, 3:30 PM

    * **Default Alignment**: **Right-aligned**

    * **Behavior**: You can calculate the difference between dates, add days, etc.

    ---

4. **Boolean (Logical)**: Represents TRUE or FALSE values, typically returned by logical formulas.

    * **Examples**: TRUE, FALSE

    * **Behavior**: Used in conditions and decision-making formulas.

---

### How to Identify Data Types:

* Look at the alignment (Left=Text, Right=Numbers/Dates).

* Look at the Number Format dropdown in the Home tab.

* Look for the **green triangle** in the top-left corner of a cell—this indicates a potential data type mismatch (e.g., a number stored as text).
---

### Importance and Real-World Use Cases

#### Why Data Types Matter:


* **Accurate Calculations**: If a "number" is stored as text, `SUM` will ignore it, leading to incorrect totals.

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


1. In cell A1, type `'1000`.

2. In cell A2, type `2000`.

3. In cell A3, write `=SUM(A1:A2)`. Note the result (it only sums 2000).

4. Convert A1 to a number. Note the sum now becomes 3000.

---

## Topic 2: Copy, Paste, and Fill Options


### Concept Explanation

#### The Power of Copy, Paste, and Fill

Copying and pasting in Excel is far more powerful than in Word. You don't just paste text; you **paste formulas**, **values**, **formats**, and even **transpose** data (switching rows to columns).

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


* **Reporting**: You calculate totals with formulas, but need to send only the final numbers to a client. You use **Paste Values** to remove the formulas.

* **Data Restructuring**: You have data in rows but need it in columns for a chart. You **use Transpose**.

* **Budget Updates**: You need to increase all prices by 10%. You type 1.1, copy it, select the price range, and use **Paste Special → Multiply**.

* **Templates**: You copy formatting from a perfect report to a new one using **Paste Formats**.

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

2. Go to the **Data tab**.

3. Click **AZ** (Sort Ascending) or **ZA** (Sort Descending).

4. _Excel Warning_: If your data has a header row, Excel will usually detect it and ask if you want to expand the selection. **Always expand** to keep rows intact.

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

2. Go to **Data** → **Filter** (or press Ctrl+Shift+L).

3. Click the dropdown arrow in the "City" column.

4. Uncheck "Select All".

5. Check only "NY".

6. Excel now shows only rows where City is NY.


#### 2. Using Number Filters

1. Click the dropdown arrow in the "Sales" column.


2. Select **Number Filters** → **Greater Than**.

3. Enter `600`.

4. Click OK. Only sales > 600 appear.

#### 3. Clearing Filters

1. Click the **Clear** button on the Data tab (next to Filter) to show all data.


2. Or click the dropdown arrow and select "Clear Filter From [Column]".


#### 4. Filter by Selection (Right-click method)

1. Right-click a cell containing "LA".

2. Choose **Filter → Filter by Selected Cell's Value**. This instantly filters the entire list to "LA".

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

1. Press **Ctrl+F**.

2. Type "Bob" in the Find what box.

3. Click "Find Next" to cycle through each occurrence.

#### 2. Basic Replace

1. Press Ctrl+H (Find & Replace shortcut).

2. Type "Sales" in **Find what**.

3. Type "Marketing" in **Replace with**.

4. Click "Replace All" to change everything, or "Replace" to do one at a time.

#### 3. Advanced Options (Important!)

1. In Find/Replace, click **Options >>**.

2. **Within**: Choose Sheet or Workbook (search all tabs).

3. **Search**: By Rows or Columns.

4. **Look in**: Formulas, Values, or Comments.

5. **Match case**: "apple" ≠ "Apple".

6. **Match entire cell contents**: Prevents changing "123" inside "4123".


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

* **Importing from Databases**: Many systems export data as CSV (comma-separated). Text to Columns is the main way to parse this.

* **Splitting Names**: You receive a list of "Full Name" (e.g., "John Smith") but need "First Name" and "Last Name" separately for mail merging.

* **Address Extraction**: Splitting a full address into Street, City, State, Zip.

* **Parsing Log Files**: Extracting specific IDs from a combined string.

---

### Step-by-Step Demonstration

#### 1. Delimited (Comma Separated)

1. In column A, type several entries like: `John,Doe,35`, `Jane,Smith,28`.

2. Select column A.

3. Go to Data → Text to Columns.

4. Choose Delimited → Next.

5. Check Comma (uncheck others). You'll see a preview of the split.

6. Next.

7. Choose the destination (e.g., $B$1 if you want to start at B1).

8. Click Finish. Result: John | Doe | 35 in separate columns.

#### 2. Fixed Width (Untidy data)


1. Type a list of codes like `A12345B`, `C98765D`.

2. Select the column → Text to Columns → Fixed Width.

3. Click in the data preview to set breakpoints (e.g., after the 1st character, after the 6th).

4. Finish

---


### Syntax or Rules

#### Rules:


* This operation **overwrites** the original data if you don't set a new destination.

* It works on one column at a time.

* Dates split into separate columns might lose date formatting if not handled.

---

### Practice Exercises

#### Exercise 1: Splitting Full Names

1. Enter a list of full names in column A: "John Smith", "Mary Johnson", "Bob Davis".

2. Use Text to Columns (Delimited by Space) to split them into First Name and Last Name.

#### Exercise 2: Parsing Addresses

1. Enter "123 Main St, New York, NY" in A1.

2. Split it into Street, City, State using a comma delimiter.

3. Note: It will split into three columns.

---

## Topic 7: Flash Fill

### Concept Explanation

#### What is Flash Fill?

Flash Fill is Excel's "magic" tool. It detects a pattern in your data entry and automatically fills the rest of the column for you. It learns from examples.

#### How to use it:


1. In the column next to your data, type an example of what you want the output to look like.

2. Press **Ctrl+E** (or go to Data → Flash Fill).

3. Excel recognizes the pattern and fills the entire column.

#### Common Patterns Flash Fill Recognizes:

* Extracting first names from full names.

* Extracting initials.

* Combining first and last names.

* Formatting phone numbers `(123) 456-7890`.

* Changing case (lowercase, uppercase, proper case).

---

### Importance and Real-World Use Cases

* **Data Extraction**: Pulling ZIP codes out of addresses without formulas.

* **Data Standardization**: Converting "john.doe@email.com" to "John Doe" (Name formatting).

* **Reformatting**: Taking "1234567890" and making it "(123) 456-7890".

* **Data Cleanup**: Removing extra spaces or invalid characters.

---

### Step-by-Step Demonstration

#### 1. Extracting First Names


1. Column A has full names: "John Smith", "Mary Davis", "Bob Johnson".

2. In Column B1, type "John" (just the first name).

3. Select B1. Press Ctrl+E.

4. Excel fills B2 with "Mary", B3 with "Bob".

#### 2. Combining Names

1. Column A: "John", Column B: "Smith".

2. In Column C1, type "John Smith" (combine them with a space).

3. Press Ctrl+E. Excel combines all rows.

#### 3. Formatting Phone Numbers

1. Column A has "1234567890".

2. In B1, type "(123) 456-7890".

3. Press Ctrl+E. Excel formats the rest.

#### 4. Extracting Domains from Emails


1. Column A has "john@gmail.com", "mary@yahoo.com".

2. In B1, type "gmail.com".

3. Ctrl+E. Excel extracts the domain.

---

### Syntax or Rules

#### Rules for Flash Fill:

* Excel needs a clear, consistent pattern.

* If a pattern is ambiguous, Flash Fill might fail. Provide 2 examples if the pattern is complex.

* Flash Fill is NOT dynamic. If you change the source data, the Flash Fill results do not update (unlike formulas).

* Use Ctrl+E or the Data tab.

---

### Practice Exercises

#### Exercise 1: Extract Domain

1. Enter emails: `alice@company.com`, `bob@test.org`, `carol@work.net`.

2. In column B, extract the domain (e.g., `company.com`, `test.org`).

#### Exercise 2: Clean up Dates

1. Enter "2024-01-15", "2024-02-20".

2. Use Flash Fill to convert to "January 15, 2024" format (type one example).

#### Exercise 3: Extract Initials

1. Enter "John Fitzgerald Kennedy".

2. Extract the initials "JFK" using Flash Fill.

---

## Topic 8: Data Validation

### Concept Explanation

#### What is Data Validation?

Data Validation is a rule you set on a cell or range to control what kind of data can be entered. It acts as a gatekeeper, preventing users from entering incorrect or invalid data.

#### Types of Validation (Most Common):

* **Whole Number**: Only integers (e.g., between 1 and 100).

* **Decimal**: Numbers with decimals (e.g., between 0.1 and 1.0).

* **List**: Creates a dropdown menu of allowed values (e.g., "Select: Yes, No").

* **Date**: Only valid dates within a range.

* **Text Length**: Restrict the number of characters.

* **Custom**: Use a formula for advanced checks.


#### Components:

* **Settings**: The rule (e.g., "List").

* **Input Message**: A tooltip that appears when you click the cell (e.g., "Please select a department").

* **Error Alert**: A warning that appears when a user enters invalid data (can be Stop, Warning, or Information)

---

### Importance and Real-World Use Cases

* **Data Integrity**: Ensuring "Revenue" columns only accept positive numbers.

* **Standardization**: Using dropdowns for "Yes/No" prevents "Y", "N", "Yeah", "Nope" variations.

* **User Guidance**: Making it easy for others to fill out a template without mistakes.

* **Budget Control**: Limiting "Number of Employees" to whole numbers.

---

### Step-by-Step Demonstration

#### 1. Creating a Dropdown List (Most Popular Use)

1. Select the cells where you want the dropdown (e.g., B2:B10).

2. Go to **Data** → **Data Validation**.

3. **Under Allow**: select **List**.

4. In the Source: box, type your options separated by commas: `Yes, No` or refer to a range like `=$A$1:$A$3`.

5. Click OK. Now users can only choose from the dropdown.

#### 2. Limiting Numbers (Whole Number)

1. Select a range for "Age".

2. Data Validation → Allow: **Whole Number**.

3. Data: **between**.

4. Minimum: `18`, Maximum: `65`.

5. Go to **Error Alert tab**.

6. Title: "Invalid Age". Error message: "Age must be between 18 and 65."

7. Click OK. If a user types `70`, they get a warning.

#### 3. Restricting Text Length (e.g., 10-digit phone numbers)

1. Data Validation → Allow: Text Length.

2. Data: equal to.

3. Length: `10`.

#### 4. Input Message (Guidance)

1. In Data Validation, go to Input Message tab.

2. Title: "Enter State".

3. Input Message: "Please use the two-letter state abbreviation (e.g., NY, CA)."



---

### Syntax or Rules

#### Rules:


* Users can still copy/paste invalid data into a validated cell (pasting overrides validation).

* You can clear validation (Data → Data Validation → Clear All).

* You can find all cells with validation: Home → Find & Select → Data Validation.

### Practice Exercises

#### Exercise 1: Create a Dropdown

1. Create a list of cell values: A1:A3 = "Apple", "Banana", "Cherry".

2. In B1, create a dropdown list referencing A1:A3.

3. Now change the dropdown to use a comma-separated list instead (without referencing a range).

#### Exercise 2: Number Validation

1. Set a rule on a cell so it only accepts values between 0 and 100.

2. Test it by entering 50 (works), and 150 (error).

3. Add a custom error message.

#### Exercise 3: Date Validation

1. Allow only dates in January 2024 to be entered.

2. _Hint_: Allow Date → between `1/1/2024` and `1/31/2024`.

---

## Mini Quiz: Phase 2 - Working with Data

#### Section 1: Data Types

1. How can you tell if a number is stored as text just by looking at it?

2. What is the fastest way to convert a column of numbers stored as text to real numbers?

3. What alignment do dates typically have in Excel?

#### Section 2: Copy, Paste, Fill

4. What is the shortcut for Paste Special Values?

5. Why would you use Paste Values instead of a normal Paste?

6. How do you fill a series of numbers 1, 3, 5, 7 using the Fill Handle?

#### Section 3: Sorting & Filtering

7. What is the risk of sorting only one column in a table without expanding the selection?

8. What does a filter do to rows that don't meet the criteria?

9. How do you filter to show only records where sales are greater than $5,000?

#### Section 4: Find, Replace, Text to Columns

10. What keyboard shortcut opens Find & Replace directly (Replace tab)?

11. When splitting "John|Doe|35" into columns, what delimiter would you use?

12. What is the difference between 'Delimited' and 'Fixed Width' in Text to Columns?

#### Section 5: Flash Fill & Data Validation

13. What is the keyboard shortcut for Flash Fill?

14. Why is Flash Fill not ideal for data that changes regularly?

15. How do you create a dropdown list in Excel?

---

## Common Mistakes and Best Practices

1. **Sorting Only One Column**: Mistake: Selecting a single column and sorting it, breaking all row relationships. Solution: Always select the entire table or let Excel expand the selection.

2. **Hardcoding Values via Paste**: Mistake: Using regular Paste over formulas and losing cell references. Solution: Use Paste Values when you intend to remove formulas, but be aware you lose dynamic updates.

3. **Ignoring the Green Triangle**: Mistake: Assuming numbers with a green triangle are fine, then wondering why SUM returns 0. Solution: Always convert text-looking numbers to actual numbers.

4. **Overwriting Data with Text to Columns**: Mistake: Running Text to Columns on a column and not changing the destination, losing the original data. Solution: Always select a destination column if you want to keep the original.

5. **Over-relying on Flash Fill for Dynamic Data**: Mistake: Using Flash Fill to combine data, then changing the source data and wondering why the output doesn't update. Solution: Use formulas (`&` or `CONCATENATE`) if the data changes frequently.

6. **Forgetting to Clear Filters**: Mistake: Filtering data, printing it, and then wondering why half the data is missing. Solution: Always clear filters (`Ctrl+Shift+L` twice) or look for the filter icon in the header.

7. **Copying with Filters On**: Mistake: Copying data while filters are active and pasting it elsewhere, inadvertently pasting only visible rows. Solution: Be mindful of visible vs. hidden rows. (Note: Paste Special can paste only visible cells).

8. **Not Using Data Validation**: Mistake: Allowing free-text entry for "Department" leading to 20 different spellings of "Finance". Solution: Use Data Validation List dropdowns to enforce consistency.

---

## Best Practices

1. **Always Use Headers**: Before sorting or filtering, ensure your data has a clear header row. It makes selecting and sorting intuitive.

2. **Backup Before Large Operations**: Before doing a massive Find & Replace or Text to Columns, save a copy of your workbook. Mistakes are easier to revert.

3. **Use Paste Values for Final Reports**: When distributing a report, paste values over all formulas to prevent recipients from accidentally changing them, and to speed up the file.

4. **Combine Flash Fill with Formulas**: Use Flash Fill for quick, one-off extractions. Use `LEFT`, `RIGHT`, `MID`, `FIND` formulas for consistent, dynamic extractions.

5. **Validate Input Early**: Set Data Validation on templates before sending them to other people. It drastically reduces data cleaning time later.

6. **Use "Match Entire Cell Contents" in Find**: This prevents accidental replacements (e.g., replacing "1" with "2" won't turn "15" into "25").

7. **Document Your Data Cleaning**: If you use complex Text to Columns or Filters, keep a note of what you did. Future you will thank you.

---

## Interview Questions

### Beginner Level

1. **Explain the difference between sorting and filtering.**

    Sorting rearranges the rows of your data into a specific order (ascending/descending). Filtering temporarily hides rows that don't meet a condition, without changing their order.

2. **How do you convert a column of numbers stored as text?**

    Select the column, click the yellow warning diamond, and select "Convert to Number." Alternatively, use Text to Columns and click Finish.

3. **What is the Fill Handle and what can it do?**

    It's the small square at the bottom-right of a selected cell. It can copy values, fill series (like months/numbers), and copy formulas.

4. **What is Paste Special? Give an example.**

    Paste Special lets you choose what part of the copied data to paste. Example: Pasting "Values" only pastes the result of a formula, not the formula itself.

5. **Why would you use Text to Columns?**

    To split a single column of data into multiple columns based on a delimiter (like a comma) or fixed width, useful for parsing CSV files or full names.

### Intermediate Level

6. **What happens if you sort a range that has blank rows in between?**

    Excel sorts each continuous block of data separately. This is why data tables should have NO blank rows or columns within the data range.

7. **Can you filter by color? How?**

    Yes. Enable filters, click the dropdown arrow in the header, go to "Filter by Color", and choose the cell or font color.

8. **How does Flash Fill differ from Text to Columns?**

    Flash Fill is pattern-based and doesn't use a delimiter; it works in the adjacent column. Text to Columns splits within the same column or to the right and requires a clear delimiter or fixed widths. Flash Fill is temporary; Text to Columns permanently splits the data.

9. **How can you find all cells with Data Validation in a workbook?**

    Go to Home → Find & Select → Data Validation. Excel will select all cells that have validation rules applied.

10. **If you use Flash Fill and change the original data, does the Flash Fill output update? Why?**

    No. Flash Fill is a static operation, not a dynamic formula. It only copies the values at the moment it's run. You need to re-run Flash Fill (Ctrl+E) or use formulas for dynamic updates.


---

## Assignment and Project Work: Data Cleaning & Organization Project

#### Scenario: 

> Messy Sales Data Cleanup

You have been provided with a messy export from an ancient sales system. Your job is to clean it up and organize it so the sales director can analyze it.

**Download/Setup**: Copy this "messy" data into a new Excel workbook starting at A1.


| Full Name     | Phone         | Email            | Region/State | Sales_Amount |
|---------------|---------------|------------------|--------------|--------------|
| John_Doe      | 1234567890    | john@abc.com     | NY           | 1500.5       |
| Jane_Smith    | 987-654-3210  | jane@xyz.org     | CA           | 2500.75      |
| Bob Johnson   | 5551234567    | bob@abc.com      | ny           | 1000.00      |
| Mary_Davis    | 212-555-1212  | mary@test.net    | CA           | 3000         |
| (Empty)       | 555-9876      |                  | NY           | 500          |

### Tasks to Perform:

#### 1. Data Types & Cleaning:

* Fix the `Sales_Amount` column. Ensure all numbers are proper numbers and right-aligned. (Hint: check for text formatting).

* Standardize the `Region/State` column. Ensure all are uppercase ("NY", "CA").

#### 2. Text to Columns:

* Split the `Full Name` column into two columns: `First Name` and `Last Name`.

* Hint: The delimiter is either an underscore (`_`) or a space ( ). You might need to handle both.

#### 3. Flash Fill:

* In a new column called `Domain`, extract the domain from the `Email` column (e.g., `abc.com`).

#### 4. Find & Replace:

* There is a blank cell in `Full Name`. Replace it with "Unknown".

#### 5. Data Validation:

* Apply Data Validation to the `Region/State` column to only allow "NY" or "CA" from a dropdown list.

#### 6. Sorting:

* Sort the entire list by `State` (A-Z) and then by `Sales_Amount` (Largest to Smallest).

#### 7. Filtering:

* Apply a filter to show only records from "NY".

* Copy the visible filtered rows to a new sheet named "NY_Sales".

#### Deliverables:

* A clean workbook with the original data (cleaned) on Sheet1.

* A new sheet named "NY_Sales" containing only NY records.

* All formulas (if any) and validations working correctly.

---

## Summary and Revision Notes

### Phase 2: Key Concepts - Quick Revision

#### 1. Data Types

* **Text (Left) vs Numbers/Dates (Right).**

* Green triangle = Number stored as Text → Convert!

* Use Format Cells to change display, not the underlying type.

#### 2. Copy, Paste, Fill

* **Ctrl+C / Ctrl+V.**

* **Paste Values (123)** = Remove formulas, keep numbers.

* **Paste Transpose** = Flip rows to columns.

* **Fill Handle**: Drag to copy; hold Ctrl while dragging to increment numbers.

#### 3. Sorting & Filtering

* **Sort**: Rearranges data. Always select the whole table.

* **Filter**: Hides data. Use Ctrl+Shift+L to toggle.

* **Number Filters**: Greater than, Top 10, etc.

#### 4. Find, Replace, Text to Columns

* **Ctrl+H** for Replace.

* **Text to Columns**: Split data by Comma, Tab, or Fixed Width.

* **Destination** is crucial—don't overwrite your source!

#### 5. Flash Fill

* **Ctrl+E** to invoke.

* Extracts/combines data based on patterns.

* **Static** (does not update automatically).

#### 6. Data Validation

* Restricts input to prevent errors.

* **List = Dropdown** menus.

* Can set Input Messages and Error Alerts.

### Best Practices Checklist

* Always ensure numbers are truly numbers (right-aligned).

* Use Paste Values when distributing final reports.

* Never sort a single column by itself.

* Clear filters before saving/closing to avoid confusion.

* Use Data Validation dropdowns to standardize entries.

* Prefer formulas over Flash Fill for dynamic/live data.

### Your Learning Journey Continues

You've mastered the essential data manipulation tools! You can now take raw, messy data and transform it into organized, structured, and validated information.

#### What You've Learned:

* Managing Data Types for accuracy.

* Efficient Copy, Paste, and Fill techniques.

* Sorting and Filtering for analysis.

* Cleaning data with Find/Replace and Text to Columns.

* Automating patterns with Flash Fill.

* Controlling input with Data Validation.

#### Ready for Phase 3: Functions

* String Functions
* Numeric Functions
* Date Functions
* Aggregate Functions
    * COUNT
    * SUM
    * AVG
    * MIN
    * MAX

#### Before Moving On:

1. Complete all practice exercises.

2. Complete the Data Cleaning Project.

3. Test yourself with the mini quiz.

4. Review any areas that feel unclear.

5. Practice with your own messy data!

----


<br/><br/><br/>
<center> <b>Happy Learning! 😊</b> </center>