## Excel 365 new features

- Dynamic array
- LAMBDA function

## Voltage drop table calculation

The VDrop.xlsx spreadsheet generates a voltage drop table for a typical motor circuit. The voltage drop table lists the maximum allowable cable length for the motor circuit based on various design parameters:

- Motor size
- Cable size
- System voltage
- Maximum voltage drop allowed
- etc...

Note, the purpose of this exercise is to utilize the new Excel 365 features and not necessary to implement all the features of a voltage drop table.

## Excel 365 dynamic array

Cable data and motor data are stored as Excel tables. Based on the cable and load design parameters, data are retrieved from the cable and motor tables to populate the the voltage drop table using Excel dynamic array functions. The maximum length calculations also use the dynamic array feature of Excel.

The dynamic array feature in Excel allows a function formula entered in a **single cell** to populate data across multiple adjacent rows and/or columns. If a formula only return a single value, then such formula must be duplicated across all the cells of these rows and/or columns. It is easier to maintain a single formula than multiple formulas. Many existing and new functions in Excel 365 accept array input parameters and return an array output.

## Excel 365 LAMBDA function

A custom function in Excel can be defined as a VBA macro or as a LAMBDA function. A LAMBDA function is defined as an Excel Name. LAMBDA function has inputs parameters and the calculations are performed using other Excel functions including other LAMBDA functions. LAMBDA function can accept array input and output array of values.

The VDrop.xlsx spreadsheet defines several LAMBDA functions:

- fxEs
- fxVd
- fxA
- fxB
- fxC
- fxL
- fxMM
- fxMTT

Note, fxMTT calculate maximum cable lengths in the voltage drop table. It has a dependency on all other LAMBDA functions.