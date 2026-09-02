Day 9 - Lookup and Reference Mastery

Topics Covered

1. INDEX and MATCH

Learned how **INDEX** and **MATCH** work together to perform flexible lookups in Excel.

INDEX – The Selector

Definition:
INDEX returns the value from a specific position within a range based on a row number and/or column number.

Why is INDEX important?**

* Can retrieve data from any position in a table.
* More flexible than traditional lookup methods.
* Can work with both rows and columns.
* Commonly used with MATCH for dynamic lookups.

Example:

excel
=INDEX(C2:C10,5)

This returns the 5th value from the range C2:C10.

MATCH – The Positional Finder

Definition:
MATCH searches for a specific value in a range and returns its position.

hy is MATCH important?

* Finds the position of a value dynamically.
* Can be combined with INDEX for flexible lookups.
* Useful when the lookup column is not the first column.

Example:

excel
=MATCH("Laptop",A2:A10,0)


This returns the position of "Laptop" within the range.

INDEX + MATCH

INDEX and MATCH can be combined to create a dynamic lookup.

excel
=INDEX(C2:C10,MATCH("Laptop",A2:A10,0))


This finds "Laptop" in column A and returns the corresponding value from column C.

 Common Use Cases

* Finding employee salaries using Employee ID
* Finding product prices using Product ID
* Retrieving customer information
* Looking up department or location
* Matching data between different tables
* Creating dynamic reports and dashboards

2. Dynamic Lookups

INDIRECT

Definition:
INDIRECT converts a text string into a cell or range reference.

Example:

excel
=INDIRECT("B5")


This returns the value from cell B5.

Common use:
Creating dynamic references based on text or cell values.

OFFSET

Definition:
OFFSET returns a reference that is a specified number of rows and columns away from a starting cell.
Example:

excel
=OFFSET(A1,2,1)


This refers to the cell that is 2 rows down and 1 column to the right of A1.

Common use:
Creating dynamic ranges and flexible calculations.

Named Ranges

Definition:
Named Ranges allow a meaningful name to be assigned to a cell or range instead of using cell references.

Example:

Instead of:

excel
=SUM(B2:B20)


A range can be named **Sales** and used as:

excel
=SUM(Sales)


Benefits:

* Makes formulas easier to understand.
* Improves readability.
* Makes large worksheets easier to manage.
* Useful for dashboards and dynamic reports.

Skills Practiced

* INDEX
* MATCH
* INDEX + MATCH
* INDIRECT
* OFFSET
* Named Ranges
* Dynamic Lookups
* Data Retrieval
* Reference Management
