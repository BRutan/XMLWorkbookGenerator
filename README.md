# XMLWorkbookGenerator

XMLWorkbookGenerator is an intuitive wrapper for the OpenXmlWriter class defined in the OpenXML library. 

OpenXmlWriter is a very fast, low memory overhead way of generating Excel workbooks, by writing to the underlying XML inherent to XLSX files. It unfortunately is rather complicated and mysterious to work with, given the lack of documentation. Hence the necessity for this class.

XMLWorkbookGenerator can generate a 150 MB workbook filled with simple strings in one minute and thirty seconds, using at most 75MB of RAM. This is astronomically faster than VBA, and would cause an out-of-memory issue if attempted with the default C# Excel interop library. 

XMLWorkbookGenerator is derived from FileType base class that implements useful functionality for files. It can only generate files with XLSX extension. Operations must be done in a particular order for the workbook to be generated correctly. One must first use the AddSheet(sheetName) method to set the ActiveSheet. All write operations will then be performed on this active sheet, and then locked in by calling the FinishSheet() method. After FinishSheet() has been called, one cannot write to a previously generated sheet. After writing all data to all intended sheets, one then must call the GenerateFile() method to finish writing to the file. Performing these steps out of order will throw a derived NonFatal or SemiFatal exception.    

Write operations to sheets are performed with (using zero-based indices): 

o	WriteRow(data, row#, col#): write string data to active sheet starting at initial address, in column-by-column fashion. 

o	WriteColumn(data, row#, col#): write string data to active sheet starting at initial address, in row-by-row fashion. 

o	WriteAllData(data, row#, col#): write all string data contained in 2 dimensional data structure to active sheet starting at initial address. 

# Important notes: 

o	Write operations check if provided row and column are out of bounds, and will throw a nonfatal exception if either are. The maximum (row, column) is (1048576, 16384) for XLSX workbooks. 

o	A maximum of 255 sheets can be written to the workbook. Attempting to write more will throw a SemiFatal exception. 

o	The constructor requires the full file path (folder + file name) where the workbook will be generated. It will throw an exception if the file extension is not XLSX. 

o	One cannot overwrite to previously written-to row with another call to the listed write operations, due to the nature of the underlying XML. Doing so will throw an exception, to prevent workbook corruption. 

o	Sheet names must be unique in Excel workbooks. Trying to add a sheet with a name that has already been added will throw an exception. 

o	If an Excel workbook with the same name + path exists at the path provided at the constructor, an exception will be thrown to prevent unintentional overwrites, since it would be overwritten otherwise. The relevant unit test for this class is “WorkbookGeneratorTest.cs” in CSharpObjectLibrary.
