# Difference Between SUBTOTAL and AGGREGATE Functions in Excel



##### Overview



Both SUBTOTAL and AGGREGATE are specialized functions in Excel used to compute aggregate values like SUM, AVERAGE, COUNT, and more. However, AGGREGATE is a newer, more powerful function that expands upon SUBTOTAL's capabilities, offering greater flexibility and control.



##### Supported Operations



SUBTOTAL: Supports 11 basic functions (such as SUM, AVERAGE, COUNT, MIN, MAX, etc.).



AGGREGATE: Supports 19 functions, covering all SUBTOTAL options and adding advanced statistics like MEDIAN, MODE, LARGE, SMALL, and various percentiles



##### Handling Hidden Rows and Errors



SUBTOTAL:



* Can ignore values in hidden rows only if specific function numbers (101–111) are used.
* Automatically ignores other SUBTOTAL functions nested within the calculation.



AGGREGATE:



* Allows you to choose how to handle hidden rows, errors, and nested SUBTOTAL/AGGREGATE results via its options argument.
* Can simultaneously ignore hidden rows, error values, and/or nested aggregate functions for even more precise calculations



##### Syntax and Customization



SUBTOTAL:



* =SUBTOTAL(function\_num, ref1, \[ref2], ...)
* function\_num determines the operation and whether to ignore hidden rows.



AGGREGATE:



* =AGGREGATE(function\_num, options, ref1, \[ref2], …)
* function\_num selects the calculation.
* options lets you customize which items to ignore (e.g., errors, hidden, nested functions)
