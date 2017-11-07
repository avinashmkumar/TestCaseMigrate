# TestCaseMigrate
Tool created using VBScript in MS Excel which converted existing test cases written using MS Excel in multiple legacy formats to HP ALM compatible format. 

The code scans specific cell containing multiple Test Steps in each row and matches it with a regular expression, split them at new line and writes to individual rows. Step numbers are generated dynamically for each such test case within a test suite. Upload location, author name and starting data row is accepted as parameters through a dialog.

