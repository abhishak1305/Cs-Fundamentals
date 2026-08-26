# Excel
> Excel + PowerPoint + Word + Outlook + Teams

------------------------------------------------------------------------

## Table of Contents

1.  [How to Use This Guide](#how-to-use-this-guide)
2.  [Excel Basics](#excel-basics)
3.  [Keyboard Shortcuts](#keyboard-shortcuts)
4.  [Basic Formulas](#basic-formulas)
5.  [IF, AND, OR, and NOT](#if-and-or-and-not)
6.  [COUNTIF, SUMIF, COUNTIFS, and
    SUMIFS](#countif-sumif-countifs-and-sumifs)
7.  [VLOOKUP and XLOOKUP](#vlookup-and-xlookup)
8.  [Text Functions](#text-functions)
9.  [Rounding Functions](#rounding-functions)
10. [Date and Time Functions](#date-and-time-functions)
11. [Cell References](#cell-references)
12. [Sorting and Filtering](#sorting-and-filtering)
13. [Conditional Formatting](#conditional-formatting)
14. [Data Validation](#data-validation)
15. [Remove Duplicates](#remove-duplicates)
16. [Freeze Panes](#freeze-panes)
17. [PivotTables](#pivottables)
18. [Charts](#charts)
19. [Common Excel File Terms](#common-excel-file-terms)
20. [High-Priority Revision List](#high-priority-revision-list)
21. [Practice Questions](#practice-questions)
22. [Quick Revision Sheet](#quick-revision-sheet)

------------------------------------------------------------------------

## How to Use This Guide

If you are preparing for an OA and have limited time, use this order:

1.  Learn the keyboard shortcuts.
2.  Understand basic formulas.
3.  Practice `IF`, `COUNTIF`, `SUMIF`, `COUNTIFS`, and `SUMIFS`.
4.  Learn `VLOOKUP`, `XLOOKUP`, and cell references.
5.  Practice text and date functions.
6.  Review Sort, Filter, PivotTables, Conditional Formatting, and Data
    Validation.
7.  Solve the practice questions without looking at the answers.

------------------------------------------------------------------------

# Excel Basics

## Workbook

A **workbook** is the complete Excel file.

Example:

``` text
employee_data.xlsx
```

A workbook can contain multiple worksheets.

## Worksheet

A **worksheet** is an individual sheet inside a workbook.

Example:

``` text
Sheet1
Sheet2
Sheet3
```

## Cell

A **cell** is the intersection of a row and a column.

Example:

``` text
B5
```

means column `B`, row `5`.

## Range

A **range** is a group of cells.

Example:

``` text
A1:C10
```

This represents cells from A1 through C10.

## Formula

An Excel formula normally begins with `=`.

Example:

``` excel
=A1+B1
```

------------------------------------------------------------------------

# Keyboard Shortcuts

These are high-value shortcuts to memorize.

  Shortcut                 Action
  ------------------------ ---------------------------------------------------------------
  `Ctrl + C`               Copy
  `Ctrl + X`               Cut
  `Ctrl + V`               Paste
  `Ctrl + Z`               Undo
  `Ctrl + Y`               Redo
  `Ctrl + S`               Save
  `Ctrl + F`               Find
  `Ctrl + H`               Find and Replace
  `Ctrl + A`               Select all
  `Ctrl + B`               Bold
  `Ctrl + I`               Italic
  `Ctrl + U`               Underline
  `Ctrl + P`               Print
  `Ctrl + N`               New workbook
  `Ctrl + O`               Open workbook
  `Ctrl + 1`               Open Format Cells
  `Ctrl + ;`               Insert current date
  `Ctrl + Shift + ;`       Insert current time
  `F2`                     Edit active cell
  `F4`                     Repeat action / toggle reference type while editing a formula
  `Alt + =`                AutoSum
  `Ctrl + Arrow`           Jump to the edge of a data region
  `Ctrl + Shift + Arrow`   Select data to the edge of a region
  `Shift + Space`          Select entire row
  `Ctrl + Space`           Select entire column

### Example: F4 and Absolute References

Suppose you enter:

``` excel
=A1*B1
```

While editing a reference, pressing `F4` can cycle through reference
styles such as:

``` text
A1
$A$1
A$1
$A1
```

------------------------------------------------------------------------

# Basic Formulas

## SUM

Adds numbers.

``` excel
=SUM(A1:A10)
```

Example:

If A1:A3 contains:

``` text
10
20
30
```

then:

``` excel
=SUM(A1:A3)
```

returns:

``` text
60
```

------------------------------------------------------------------------

## AVERAGE

Calculates the arithmetic average.

``` excel
=AVERAGE(A1:A10)
```

Example:

``` text
10, 20, 30
```

Average:

``` text
20
```

------------------------------------------------------------------------

## MAX

Returns the largest value.

``` excel
=MAX(A1:A20)
```

------------------------------------------------------------------------

## MIN

Returns the smallest value.

``` excel
=MIN(A1:A20)
```

------------------------------------------------------------------------

## COUNT

Counts cells containing numbers.

``` excel
=COUNT(A1:A10)
```

Example:

  Cell   Value
  ------ -------
  A1     10
  A2     20
  A3     Hello
  A4     30

``` excel
=COUNT(A1:A4)
```

Result:

``` text
3
```

`COUNT` does not count the text `Hello`.

------------------------------------------------------------------------

## COUNTA

Counts non-empty cells.

``` excel
=COUNTA(A1:A10)
```

Unlike `COUNT`, `COUNTA` also counts text.

------------------------------------------------------------------------

## COUNTBLANK

Counts empty cells.

``` excel
=COUNTBLANK(A1:A10)
```

------------------------------------------------------------------------

# IF, AND, OR, and NOT

## IF

`IF` checks a condition and returns one result if it is true and another
if it is false.

### Syntax

``` excel
=IF(condition, value_if_true, value_if_false)
```

### Example

``` excel
=IF(A1>=40,"Pass","Fail")
```

If A1 is 40 or greater:

``` text
Pass
```

Otherwise:

``` text
Fail
```

------------------------------------------------------------------------

## Nested IF

You can put an IF function inside another IF.

``` excel
=IF(A1>=90,"A",IF(A1>=75,"B",IF(A1>=60,"C","D")))
```

Concept:

``` text
90+  -> A
75-89 -> B
60-74 -> C
Below 60 -> D
```

------------------------------------------------------------------------

## AND

`AND` returns TRUE only when all conditions are TRUE.

``` excel
=AND(A1>50,B1>50)
```

Both conditions must be true.

### Common combination

``` excel
=IF(AND(A1>=40,B1>=40),"Pass","Fail")
```

------------------------------------------------------------------------

## OR

`OR` returns TRUE when at least one condition is TRUE.

``` excel
=OR(A1>50,B1>50)
```

### Example

``` excel
=IF(OR(A1>=40,B1>=40),"Eligible","Not Eligible")
```

------------------------------------------------------------------------

## NOT

`NOT` reverses TRUE and FALSE.

``` excel
=NOT(A1>50)
```

If `A1>50` is TRUE, `NOT` returns FALSE.

------------------------------------------------------------------------

# COUNTIF, SUMIF, COUNTIFS, and SUMIFS

## COUNTIF

Counts cells that satisfy one condition.

### Syntax

``` excel
=COUNTIF(range, criteria)
```

### Example

``` excel
=COUNTIF(A1:A10,"Pass")
```

Counts how many cells contain `Pass`.

### Numeric condition

``` excel
=COUNTIF(B1:B10,">=40")
```

Counts values greater than or equal to 40.

------------------------------------------------------------------------

## SUMIF

Adds values based on one condition.

### Syntax

``` excel
=SUMIF(range, criteria, sum_range)
```

### Example

Suppose:

  A            B
  -------- -----
  Apple      100
  Orange     200
  Apple      150

Formula:

``` excel
=SUMIF(A1:A3,"Apple",B1:B3)
```

Result:

``` text
250
```

------------------------------------------------------------------------

## COUNTIFS

Counts cells/rows satisfying multiple conditions.

``` excel
=COUNTIFS(A1:A10,"Sales",B1:B10,">50000")
```

This counts rows where:

-   Column A = `Sales`
-   Column B \> `50000`

------------------------------------------------------------------------

## SUMIFS

Adds values using multiple conditions.

``` excel
=SUMIFS(C1:C10,A1:A10,"Sales",B1:B10,">50000")
```

Here:

-   `C1:C10` = values to add
-   `A1:A10` = first criteria range
-   `"Sales"` = first condition
-   `B1:B10` = second criteria range
-   `">50000"` = second condition

------------------------------------------------------------------------

# VLOOKUP and XLOOKUP

Lookup functions are important for Excel assessments and interviews.

## VLOOKUP

`VLOOKUP` means **Vertical Lookup**.

### Syntax

``` excel
=VLOOKUP(lookup_value, table_array, col_index_num, FALSE)
```

### Example

``` excel
=VLOOKUP(A2,$F$2:$H$10,3,FALSE)
```

This:

1.  Takes the value in `A2`.
2.  Searches for it in the first column of `F2:H10`.
3.  Returns the value from the third column.
4.  Uses `FALSE` for an exact match.

### Important VLOOKUP rule

The lookup value must be in the **first column of the selected table
array**.

------------------------------------------------------------------------

## XLOOKUP

`XLOOKUP` is a newer and more flexible lookup function.

### Syntax

``` excel
=XLOOKUP(lookup_value, lookup_array, return_array)
```

### Example

``` excel
=XLOOKUP(A2,F2:F10,H2:H10)
```

It searches `F2:F10` for A2 and returns the corresponding value from
`H2:H10`.

### VLOOKUP vs XLOOKUP

  -----------------------------------------------------------------------
  VLOOKUP                             XLOOKUP
  ----------------------------------- -----------------------------------
  Older lookup function               Newer lookup function

  Uses a column number                Uses a return range

  Lookup value must be in first       Can look in different directions
  column                              

  Less flexible                       More flexible
  -----------------------------------------------------------------------

For OA preparation, understand both the **purpose and basic syntax**.

------------------------------------------------------------------------

# Text Functions

## LEFT

Returns characters from the left side of a text string.

``` excel
=LEFT(A1,3)
```

Example:

``` text
ACCENTURE
```

Result:

``` text
ACC
```

------------------------------------------------------------------------

## RIGHT

Returns characters from the right side.

``` excel
=RIGHT(A1,3)
```

Example:

``` text
ACCENTURE
```

Result:

``` text
URE
```

------------------------------------------------------------------------

## MID

Returns characters from the middle.

### Syntax

``` excel
=MID(text, start_num, num_chars)
```

Example:

``` excel
=MID(A1,2,4)
```

------------------------------------------------------------------------

## LEN

Returns the number of characters.

``` excel
=LEN(A1)
```

Example:

``` text
HELLO
```

Result:

``` text
5
```

------------------------------------------------------------------------

## UPPER

Converts text to uppercase.

``` excel
=UPPER(A1)
```

Example:

``` text
hello
```

becomes:

``` text
HELLO
```

------------------------------------------------------------------------

## LOWER

Converts text to lowercase.

``` excel
=LOWER(A1)
```

------------------------------------------------------------------------

## TRIM

Removes extra spaces from text.

``` excel
=TRIM(A1)
```

This is useful when data contains accidental leading, trailing, or
repeated spaces.

------------------------------------------------------------------------

## CONCAT

Combines text.

``` excel
=CONCAT(A1,B1)
```

Example:

``` text
A1 = John
B1 = Doe
```

Result:

``` text
JohnDoe
```

If you want a space:

``` excel
=CONCAT(A1," ",B1)
```

Result:

``` text
John Doe
```

------------------------------------------------------------------------

## TEXTJOIN

Combines multiple text values using a delimiter.

``` excel
=TEXTJOIN(", ",TRUE,A1:A5)
```

Example output:

``` text
Apple, Orange, Banana
```

------------------------------------------------------------------------

# Rounding Functions

## ROUND

Rounds a number to a specified number of decimal places.

``` excel
=ROUND(A1,2)
```

Example:

``` text
12.3456 -> 12.35
```

------------------------------------------------------------------------

## ROUNDUP

Rounds away from zero.

``` excel
=ROUNDUP(A1,2)
```

------------------------------------------------------------------------

## ROUNDDOWN

Rounds toward zero.

``` excel
=ROUNDDOWN(A1,2)
```

------------------------------------------------------------------------

# Date and Time Functions

## TODAY

Returns the current date.

``` excel
=TODAY()
```

------------------------------------------------------------------------

## NOW

Returns the current date and time.

``` excel
=NOW()
```

------------------------------------------------------------------------

## YEAR

Extracts the year from a date.

``` excel
=YEAR(A1)
```

------------------------------------------------------------------------

## MONTH

Extracts the month number.

``` excel
=MONTH(A1)
```

------------------------------------------------------------------------

## DAY

Extracts the day number.

``` excel
=DAY(A1)
```

------------------------------------------------------------------------

# Cell References

Understanding references is one of the most important Excel concepts.

## Relative Reference

``` excel
A1
```

A relative reference changes when the formula is copied.

Example:

``` excel
=A1+B1
```

Copied one row down:

``` excel
=A2+B2
```

------------------------------------------------------------------------

## Absolute Reference

``` excel
$A$1
```

An absolute reference remains fixed when copied.

Example:

``` excel
=A1*$B$1
```

When copied down:

``` text
A1 changes
$B$1 stays fixed
```

------------------------------------------------------------------------

## Mixed References

### Fixed column

``` excel
$A1
```

Column A is fixed, row can change.

### Fixed row

``` excel
A$1
```

Row 1 is fixed, column can change.

------------------------------------------------------------------------

## Reference Summary

  Reference   Column    Row
  ----------- --------- ---------
  `A1`        Changes   Changes
  `$A$1`      Fixed     Fixed
  `$A1`       Fixed     Changes
  `A$1`       Changes   Fixed

### Shortcut

While editing a formula, press:

``` text
F4
```

to cycle through reference styles.

------------------------------------------------------------------------

# Sorting and Filtering

## Sort

Sorting rearranges data.

Common options:

-   A → Z
-   Z → A
-   Smallest → Largest
-   Largest → Smallest
-   Oldest → Newest
-   Newest → Oldest

### Example

Scores:

``` text
40
90
60
70
```

Ascending sort:

``` text
40
60
70
90
```

------------------------------------------------------------------------

## Filter

Filtering displays only rows matching selected criteria.

Example:

A table contains:

``` text
IT
HR
Sales
IT
Finance
```

Filtering for `IT` displays only the IT rows.

### Difference

**Sort** changes the order.

**Filter** hides rows that don't match the criteria.

------------------------------------------------------------------------

# Conditional Formatting

Conditional Formatting automatically changes the formatting of cells
based on rules.

### Example

Suppose marks are:

``` text
35
75
42
20
90
```

A rule can highlight values below 40.

Common uses:

-   Highlight scores below a threshold.
-   Highlight duplicate values.
-   Show top/bottom values.
-   Apply color scales.
-   Display data bars.

------------------------------------------------------------------------

# Data Validation

Data Validation controls what users can enter into a cell.

### Common example

Create a dropdown:

``` text
Pending
In Progress
Completed
```

Instead of allowing arbitrary text, users select from the predefined
values.

Other uses include:

-   Restricting numbers to a range.
-   Restricting dates.
-   Creating dropdown lists.
-   Displaying input messages.

------------------------------------------------------------------------

# Remove Duplicates

The **Remove Duplicates** feature removes repeated records from a
selected data range.

Example:

``` text
A
B
A
C
B
```

After removing duplicates:

``` text
A
B
C
```

Always understand that removing duplicates changes the data, so work on
a copy when the original data must be preserved.

------------------------------------------------------------------------

# Freeze Panes

**Freeze Panes** keeps selected rows or columns visible while scrolling.

### Example

If row 1 contains column headers, freezing the top row allows you to
keep seeing:

``` text
Name | Department | Salary | ...
```

while scrolling through hundreds of records.

------------------------------------------------------------------------

# PivotTables

A **PivotTable** is used to summarize and analyze large amounts of data
quickly.

Example data:

  Employee   Department     Salary
  ---------- ------------ --------
  A          IT              40000
  B          HR              35000
  C          IT              50000
  D          Sales           45000

A PivotTable can summarize salary by department.

Possible result:

  Department     Average Salary
  ------------ ----------------
  HR                      35000
  IT                      45000
  Sales                   45000

## Main PivotTable areas

-   **Rows**
-   **Columns**
-   **Values**
-   **Filters**

### Remember

> PivotTable = summarize, group, and analyze data.

------------------------------------------------------------------------

# Charts

## Column / Bar Chart

Useful for comparing categories.

Example:

``` text
IT       50
HR       30
Sales    45
```

------------------------------------------------------------------------

## Line Chart

Best for showing trends over time.

Example:

``` text
Jan -> 100
Feb -> 120
Mar -> 150
Apr -> 140
```

------------------------------------------------------------------------

## Pie Chart

Useful for showing parts of a whole.

Example:

``` text
IT      50%
HR      30%
Sales   20%
```

Avoid using pie charts when there are too many categories.

------------------------------------------------------------------------

## Scatter Chart

Useful for examining relationships between two numerical variables.

Example:

``` text
Hours studied vs Exam score
```

------------------------------------------------------------------------

# Common Excel File Terms

  Term          Meaning
  ------------- ---------------------------------------------------------
  Workbook      Complete Excel file
  Worksheet     Individual sheet
  Cell          Intersection of row and column
  Range         Group of cells
  Formula       Expression used to calculate a result
  Function      Built-in Excel operation
  Row           Horizontal set of cells
  Column        Vertical set of cells
  Formula Bar   Area used to view/edit cell contents
  Name Box      Displays the address/name of the selected cell or range

------------------------------------------------------------------------

# High-Priority Revision List

If you have limited preparation time, prioritize these.

## Priority 1 --- Must Know

### Shortcuts

``` text
Ctrl + C
Ctrl + X
Ctrl + V
Ctrl + Z
Ctrl + Y
Ctrl + S
Ctrl + F
Ctrl + H
Ctrl + A
Ctrl + 1
F2
F4
Alt + =
Ctrl + Arrow
Ctrl + Shift + Arrow
```

### Formulas

``` text
SUM
AVERAGE
MAX
MIN
COUNT
COUNTA
IF
COUNTIF
SUMIF
COUNTIFS
SUMIFS
VLOOKUP
XLOOKUP
```

### Concepts

``` text
Relative reference
Absolute reference
Mixed reference
Sort
Filter
PivotTable
Conditional Formatting
Freeze Panes
Data Validation
Remove Duplicates
```

------------------------------------------------------------------------

## Priority 2 --- Important

``` text
LEFT
RIGHT
MID
LEN
TRIM
UPPER
LOWER
CONCAT
TEXTJOIN
ROUND
ROUNDUP
ROUNDDOWN
TODAY
NOW
YEAR
MONTH
DAY
AND
OR
NOT
```

------------------------------------------------------------------------

# Practice Questions

Try answering these without looking at the answers.

## Question 1

Which shortcut opens the **Format Cells** dialog?

A. `Ctrl + F`

B. `Ctrl + 1`

C. `Ctrl + 2`

D. `Alt + 1`

```{=html}
<details>
```
```{=html}
<summary>
```
Answer
```{=html}
</summary>
```
**B --- Ctrl + 1**

```{=html}
</details>
```

------------------------------------------------------------------------

## Question 2

Which function calculates the average of A1 through A10?

A.  

``` excel
=SUM(A1:A10)
```

B.  

``` excel
=AVERAGE(A1:A10)
```

C.  

``` excel
=COUNT(A1:A10)
```

D.  

``` excel
=AVG(A1:A10)
```

```{=html}
<details>
```
```{=html}
<summary>
```
Answer
```{=html}
</summary>
```
**B --- `=AVERAGE(A1:A10)`**

```{=html}
</details>
```

------------------------------------------------------------------------

## Question 3

Which function counts cells containing numbers?

A. `COUNTA`

B. `COUNT`

C. `COUNTBLANK`

D. `SUM`

```{=html}
<details>
```
```{=html}
<summary>
```
Answer
```{=html}
</summary>
```
**B --- COUNT**

```{=html}
</details>
```

------------------------------------------------------------------------

## Question 4

What does this formula do?

``` excel
=IF(A1>=40,"Pass","Fail")
```

```{=html}
<details>
```
```{=html}
<summary>
```
Answer
```{=html}
</summary>
```
It returns `Pass` when A1 is greater than or equal to 40. Otherwise, it
returns `Fail`.

```{=html}
</details>
```

------------------------------------------------------------------------

## Question 5

What does `$A$1` represent?

A. Relative reference

B. Absolute reference

C. Mixed reference

D. Range

```{=html}
<details>
```
```{=html}
<summary>
```
Answer
```{=html}
</summary>
```
**B --- Absolute reference**

```{=html}
</details>
```

------------------------------------------------------------------------

## Question 6

Which feature is best for summarizing a large dataset by category?

A. WordArt

B. PivotTable

C. Mail Merge

D. Spell Check

```{=html}
<details>
```
```{=html}
<summary>
```
Answer
```{=html}
</summary>
```
**B --- PivotTable**

```{=html}
</details>
```

------------------------------------------------------------------------

## Question 7

Which function can look up a value in a table and return a corresponding
value?

A. `SUM`

B. `LOOK`

C. `VLOOKUP`

D. `COUNT`

```{=html}
<details>
```
```{=html}
<summary>
```
Answer
```{=html}
</summary>
```
**C --- VLOOKUP**

```{=html}
</details>
```

------------------------------------------------------------------------

## Question 8

Which feature can create a dropdown list in a cell?

A. Data Validation

B. Conditional Formatting

C. Freeze Panes

D. PivotTable

```{=html}
<details>
```
```{=html}
<summary>
```
Answer
```{=html}
</summary>
```
**A --- Data Validation**

```{=html}
</details>
```

------------------------------------------------------------------------

# Quick Revision Sheet

## Shortcuts

``` text
Copy              Ctrl + C
Cut               Ctrl + X
Paste             Ctrl + V
Undo              Ctrl + Z
Redo              Ctrl + Y
Save              Ctrl + S
Find              Ctrl + F
Replace           Ctrl + H
Select All        Ctrl + A
Format Cells      Ctrl + 1
Edit Cell         F2
Reference Toggle  F4
AutoSum           Alt + =
Current Date      Ctrl + ;
Current Time      Ctrl + Shift + ;
Select Row        Shift + Space
Select Column     Ctrl + Space
```

## Core Formulas

``` excel
=SUM(A1:A10)
=AVERAGE(A1:A10)
=MAX(A1:A10)
=MIN(A1:A10)
=COUNT(A1:A10)
=COUNTA(A1:A10)
=COUNTBLANK(A1:A10)

=IF(A1>=40,"Pass","Fail")

=COUNTIF(A1:A10,"Pass")
=SUMIF(A1:A10,"Apple",B1:B10)

=COUNTIFS(A1:A10,"Sales",B1:B10,">50000")
=SUMIFS(C1:C10,A1:A10,"Sales",B1:B10,">50000")

=VLOOKUP(A2,F2:H10,3,FALSE)
=XLOOKUP(A2,F2:F10,H2:H10)

=LEFT(A1,3)
=RIGHT(A1,3)
=MID(A1,2,4)
=LEN(A1)
=TRIM(A1)
=UPPER(A1)
=LOWER(A1)
=CONCAT(A1,B1)

=ROUND(A1,2)
=TODAY()
=NOW()
=YEAR(A1)
=MONTH(A1)
=DAY(A1)

=AND(A1>50,B1>50)
=OR(A1>50,B1>50)
=NOT(A1>50)
```

## Core Features

``` text
Sort
Filter
PivotTable
Conditional Formatting
Data Validation
Remove Duplicates
Freeze Panes
Charts
```

------------------------------------------------------------------------

# Final ASE OA Preparation Strategy

For Excel questions, don't just memorize formulas. Learn to recognize
**what problem each formula solves**.

For example:

``` text
Need to add numbers?        -> SUM
Need an average?            -> AVERAGE
Need largest value?         -> MAX
Need smallest value?        -> MIN
Need to count numbers?      -> COUNT
Need to count non-empty?    -> COUNTA
Need a condition?           -> IF
Need conditional counting?  -> COUNTIF
Need conditional addition?  -> SUMIF
Need multiple conditions?   -> COUNTIFS / SUMIFS
Need to find a value?       -> VLOOKUP / XLOOKUP
Need text extraction?       -> LEFT / RIGHT / MID
Need character count?       -> LEN
Need to remove extra space? -> TRIM
Need rounding?              -> ROUND
Need today's date?          -> TODAY
Need fixed cell reference?  -> $A$1
Need summarize data?        -> PivotTable
Need show only matching rows? -> Filter
Need highlight based on rules? -> Conditional Formatting
Need a dropdown?            -> Data Validation
```

> **Best approach:** learn the concept → understand the syntax → solve
> 5--10 small examples → then move to timed MCQs.

------------------------------------------------------------------------

## Disclaimer

This repository is intended for educational and interview/assessment
preparation purposes. It is not an official Accenture preparation
material and does not guarantee the contents of any specific assessment.
