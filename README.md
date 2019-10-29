# XMLWorkbookGenerator

XMLWorkbookGenerator is an intuitive wrapper for the OpenXmlWriter class defined in the OpenXML library. 

OpenXmlWriter is a very fast, low overhead way of generating Excel workbooks, by writing to the underlying XML inherent to XLSX files. 

OpenXmlWriter unfortunately is rather complicated and mysterious to work with, given its lack of documentation. If one does not perform steps in the correct order, the entire workbook may become corrupted. 

XMLWorkbookGenerator can generate a 150 MB workbook filled with simple strings in one minute and thirty seconds, using at most 75MB of RAM, all else equal. 

This is astronomically faster than VBA, and would cause an out-of-memory issue if attempted with the default C# Excel interop library. 

XMLWorkbookGenerator is derived from FileType base class that implements useful functionality for files. That class is included in this repository. 

The relevant unit test for this class is “WorkbookGeneratorTest.cs” in CSharpObjectLibrary.

# Order of Operations:

Operations must be done in a particular order for the workbook to be generated correctly. The steps are:

1. AddSheet(sheetName) method to set the ActiveSheet. 
2. WriteRow(), WriteColumn() or WriteAllData() to performing all write operations on this active sheet.
3. FinishSheet() method to lock in all write operations to the worksheet.  
(Optional) Repeat steps 1-3 if necessary.
Final Step: GenerateFile() to generate file at stored path.

Performing these steps out of order will throw a derived NonFatal or SemiFatal exception.     

# Write Operations:

Write operations to sheets are performed with (using zero-based indices): 

o	WriteRow(data, row#, col#): write string data to active sheet starting at initial address, in column-by-column fashion. 

o	WriteColumn(data, row#, col#): write string data to active sheet starting at initial address, in row-by-row fashion. 

o	WriteAllData(data, row#, col#): write all string data contained in 2 dimensional data structure to active sheet starting at initial address.

# Important Notes:

o	After FinishSheet() has been called, one cannot write to a previously generated sheet. 

o	Write operations check if provided row and column are out of bounds, and will throw a NonFatal exception if either are. The maximum (row, column) is (1048576, 16384) for XLSX workbooks.

o	A maximum of 255 sheets can be written to the workbook. Attempting to write more will throw a SemiFatal exception.

o	The constructor requires the full file path (folder + file name) where the workbook will be generated. It will throw an exception if the file extension is not XLSX.

o	One cannot overwrite to previously written-to row with another call to the listed write operations, due to the nature of the underlying XML. Doing so will throw an exception, to prevent workbook corruption.

o	Sheet names must be unique in Excel workbooks. Trying to add a sheet with a name that has already been added will throw a NonFatal exception.

o	If an Excel workbook with the same name exists at the path provided at the constructor, an exception will be thrown to prevent unintentional overwrites, since it would be entirely overwritten otherwise.

o	It can only generate files with XLSX extension.
