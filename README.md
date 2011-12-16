# VBA ChangeLog

## Credit
The inspiration and starter code for this code was provided by:
    http://www.ozgrid.com/VBA/track-changes.htm

## How mine differs
The code on their page will provide you with a change log recording
the following fields:

 *  CELL CHANGED
 *  OLD VALUE
 *  NEW VALUE
 *  TIME OF CHANGE
 *  DATE OF CHANGE

My code will track the following fields:

 * SHEET
 * CELL
 * OLD VALUE
 * NEW VALUE
 * TIME
 * DATE
 * USER

The other main difference between the two is my version will also
track changes to multiple cells at one time.

## Limitations
There is one fairly major limitation to the code though. If you select a 
number of cells larger than the defined limit of an array you will crash 
the buffer.

## Install
To use simply place the code in the Private Module of the Workbook (ThisWorkbook).

## Contributing
Pull requests are welcome as always.

## License
This software is licensed under the Apache 2 license.

## Author(s)
Steven Oliver

