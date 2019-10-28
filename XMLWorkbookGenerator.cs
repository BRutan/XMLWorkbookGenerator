/* XMLWorkbookGenerator.cs
Description:
    * Generate workbooks using the OpenXML XMLWriter class using simplified and intuitive wrapper. The XMLWriter class should be used 
    over the higher level methods in the OpenXML library when generating workbooks due to highly efficient memory management and significant speed boost
    versus built-in interop libraries.
    In testing, a 150 MB .xlsx workbook can be created using simple strings in 1 minute, 30 seconds.
*/

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using CSharpObjectLibrary.FileTypes.Files;
using CSharpObjectLibrary.Exceptions;

namespace CSharpObjectLibrary.OpenXMLTools
{
    #region Objects
    /// <summary>
    /// Generate workbooks using the OpenXML.XMLWriter class using simplified and intuitive wrapper.
    /// </summary>
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId("CSharpObjectLibrary.XMLWorkbookGenerator")]
    public class XMLWorkbookGenerator : FileType, IDisposable
    {
        #region Static Members
        private static string _DefaultPath;
        private static HashSet<string> _ValidExtensions;
        private static Tuple<ulong, ulong> _MaxDimensions;
        #endregion
        #region Class Members
        public string ActiveSheet { get { return this._ActiveSheet; } }
        public int ActiveSheetIndex { get { return this._CurrentSheetIndex; } }
        public int NumSheets { get { return this._Worksheets.Count; } }
        public string WorkbookName { get { return this._WorkbookName; } }
        private SpreadsheetDocument _Workbook;
        private List<WorksheetPart> _Worksheets;
        private Dictionary<string, HashSet<ulong>> _ExistingSheetNamesToWrittenRows;
        private new List<OpenXmlAttribute> _Attributes;
        private OpenXmlWriter _MainWriter, _HelperWriter;
        private string _ActiveSheet;
        private string _WorkbookName;
        private int _CurrentSheetIndex;
        #endregion
        #region Constructors
        static XMLWorkbookGenerator()
        {
            XMLWorkbookGenerator._DefaultPath = "{LocalPath}";
            XMLWorkbookGenerator._MaxDimensions = new Tuple<ulong, ulong>(1048576, 16384);
            XMLWorkbookGenerator._ValidExtensions = new HashSet<string>()
            {
                "xlsx"
            };
        }
        /// <summary>
        /// Construct new workbook intended to be written to filePath. Ensure that enclosing folder for intended file
        /// exists and the extension is valid. If any issues occur then throw exception.
        /// If UseSharedStrings is set to true then will use shared strings to reduce memory used in workbook. 
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="UseSharedStrings"></param>
        public XMLWorkbookGenerator(string filePath) : base(filePath, "XMLWorkbookGenerator", new CSharpObjectLibrary.Utilities.ApplicationAttributes())
        {
            var errorMessage = new StringBuilder();
            var agg = new Exceptions.ExceptionAggregator();

            this._WorkbookName = FileType.GetFileName(filePath);
            // Ensure that the enclosing folder is valid:
            if (!System.IO.Directory.Exists(this.FolderPath))
            {
                agg.Append(new Exceptions.SemiFatals.MissingFolders(this.ConvertedPath, "XMLWorkbookGenerator()"));
            }
            // Ensure that passed file extension is valid:
            if (!XMLWorkbookGenerator.ExtensionIsValid(this.Extension))
            {
                errorMessage.Append("Workbook extension ");
                errorMessage.Append(this.Extension);
                errorMessage.Append(" is invalid. Must be one of ");
                errorMessage.Append(String.Join(", ", XMLWorkbookGenerator._ValidExtensions));
                errorMessage.Append("\n");
                agg.Append(new Exceptions.NonFatals.GenericValueErrors(errorMessage.ToString(), "XMLWorkbookGenerator()"));
                errorMessage.Clear();
            }
            // Ensure that workbook does not exist already:
            if (this.Exists())
            {
                errorMessage.Append("Workbook already exists at ");
                errorMessage.Append(this._Path);
                agg.Append(new Exceptions.SemiFatals.FailedToGenerateFiles(this._WorkbookName, this._Path, "XMLWorkbookGenerator()", errorMessage.ToString()));
                errorMessage.Clear();
            }
            // Generate exception if issues occurred:
            if (agg.HasErrors())
            {
                throw agg;
            }
            this._CurrentSheetIndex = 0;
            this._ExistingSheetNamesToWrittenRows = new Dictionary<string, HashSet<ulong>>();
            this._Workbook = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook);
            this._Workbook.AddWorkbookPart();
            this._Worksheets = new List<WorksheetPart>();
            this._MainWriter = OpenXmlWriter.Create(this._Workbook.WorkbookPart);
            // Write start tag for _Workbook:
            this._MainWriter.WriteStartElement(new Workbook());
            // Write start tag for sheets collection:
            this._MainWriter.WriteStartElement(new Sheets());
        }
        #endregion
        #region Class Methods
        /// <summary>
        /// This method is unnecessary for this object. Throw NotImplementedException().
        /// </summary>
        public override void ClearContents()
        {
            throw new NotImplementedException();
        }
        /// <summary>
        /// Reset this object to be written to its default path.
        /// </summary>
        public override void ResetPath()
        {
            this._Path = XMLWorkbookGenerator._DefaultPath;
        }
        /// <summary>
        /// Return string description of this object. 
        /// </summary>
        /// <returns></returns>
        public override string ToString(bool withInfo)
        {
            var output = new StringBuilder(this._ObjectName);
            if (withInfo)
            {
                output.Append("{FileName:");
                output.Append(this._WorkbookName);
                output.Append(",SheetCount:");
                output.Append(this._Worksheets.Count);
                output.Append("}");
            }
            return output.ToString();
        }
        #endregion
        #region Workbook Structure Mutators
        /// <summary>
        /// Add new sheet to the workbook, and set the active sheet to the newly created worksheet.
        /// </summary>
        /// <param name="sheetName"></param>
        public void AddSheet(string sheetName)
        {
            // Finish writing previous written sheet to workbook:
            if (this._HelperWriter != null)
            {
                this.FinishSheet();
            }
            // Throw exception if sheet with name already exists in workbook: 
            if (this._ExistingSheetNamesToWrittenRows.ContainsKey(sheetName))
            {
                throw new Exceptions.NonFatals.GenericValueErrors("Sheet " + sheetName + " already exists in this workbook.", "XMLWorkbookGenerator::AddSheet()");
            }
            if (this._ExistingSheetNamesToWrittenRows.Keys.Count > 255)
            {
                throw new Exceptions.NonFatals.GenericValueErrors("The maximum number of sheets (255) have been added to the workbook.", "XMLWorkbookGenerator::AddSheet()");
            }
            // Create new worksheet:
            this._Worksheets.Add(this._Workbook.WorkbookPart.AddNewPart<WorksheetPart>());
            this._CurrentSheetIndex = this._Worksheets.Count - 1;
            // Update to use new sheet name, add written row # tracker for each sheet:
            this._ActiveSheet = sheetName;
            this._ExistingSheetNamesToWrittenRows.Add(sheetName, new HashSet<ulong>());
            // Point the helper writer to the new worksheet:
            this._HelperWriter = OpenXmlWriter.Create(this._Worksheets[this._CurrentSheetIndex]);
            // Write start tag for new Worksheet:
            this._HelperWriter.WriteStartElement(new Worksheet());
            // Write start tag for Sheet Data:
            this._HelperWriter.WriteStartElement(new SheetData());
        }
        /// <summary>
        /// Finish writing to the active sheet. Will do nothing if there is no active sheet. 
        /// </summary>
        public void FinishSheet()
        {
            if (this._HelperWriter != null)
            {
                // Write end tag for Sheet Data:
                this._HelperWriter.WriteEndElement();
                // Write end tag for Worksheet:
                this._HelperWriter.WriteEndElement();
                this._HelperWriter.Close();

                this._MainWriter.WriteElement(new Sheet()
                {
                    Name = this._ActiveSheet,
                    SheetId = (uint)this._Worksheets.Count,
                    Id = this._Workbook.WorkbookPart.GetIdOfPart(this._Worksheets[this._CurrentSheetIndex]),
                    State = SheetStateValues.Visible
                });
                this._HelperWriter = null;
            }
        }
        /// <summary>
        /// Output the workbook at the stored filepath. 
        /// </summary>
        public void GenerateFile()
        {
            // Write end tag to sheets collection:
            if (this._HelperWriter != null)
            {
                // Finish writing sheet:
                this.FinishSheet();
            }
            if (this._Workbook != null)
            {
                // Write end tag to sheet data:
                this._MainWriter.WriteEndElement();
                // Write end tag to workbook:
                this._MainWriter.WriteEndElement();
                // Close Writer:
                this._MainWriter.Close();
                this._Workbook.Dispose();
                this._MainWriter = null;
                this._Workbook = null;
            }
        }
        #endregion
        #region Data Writing
        /// <summary>
        /// <para>Write to all cells contained in the data collection to current sheet in column-by-column fashion </para>
        /// <para>(i.e. will fill the column starting at passed coordinates). Note that you must have created at least one sheet</para>
        /// <para>before you can write to the workbook and you can only write to the active sheet. </para>
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <param name="data_in"></param>
        public void WriteRow<T>(T data_in, ulong row = 0, ulong col = 0, RangeFormatting? formats = null) where T : IEnumerable<string>
        {
            List<string> data = data_in.ToList();

            // Switch to 1-based indices:
            row++;
            col++;

            // Throw exception if maximum coordinates will be out of bounds:
            if (OutOfBounds(row, col + (ulong)data.Count))
            {
                throw new Exceptions.NonFatals.GenericValueErrors(String.Format("R{0}C{1} is out of bounds for this workbook type.", row, col + (ulong)data.Count), "XMLWorkbookGenerator::WriteColumn()");
            }
            // Throw exception if row has already been written to:
            if (this._ExistingSheetNamesToWrittenRows[this._ActiveSheet].Contains(row))
            {
                var message = String.Format("Row {0} has already been written to.", row);
                throw new Exceptions.SemiFatals.FailedToCreateSheets(this._ActiveSheet, "XMLWorkbookGenerator::WriteRow()", this._WorkbookName, message);
            }
            if (data.Count > 0)
            {
                // Update the written row tracker:
                this._ExistingSheetNamesToWrittenRows[this._ActiveSheet].Add(row);
                // Write start tag for row:
                this._Attributes = new List<OpenXmlAttribute>();
                this._Attributes.Add(new OpenXmlAttribute("r", null, row.ToString()));
                this._HelperWriter.WriteStartElement(new Row(), this._Attributes);
                // Write all elements in container in column-by-column fashion:
                foreach (var elem in data)
                {
                    this.WriteToCell(row, col, elem);
                    col++;
                }
                // Write end tag for row:
                this._HelperWriter.WriteEndElement();
            }
        }
        /// <summary>
        /// Write to all cells contained in the data collection in row-by-row fashion (i.e. will fill the column) starting
        /// at passed coordinates. 
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <param name="data_in"></param>
        public void WriteColumn<T>(T data_in, ulong row = 0, ulong col = 0, RangeFormatting? formats = null) where T : IEnumerable<string>
        {
            List<string> data = data_in.ToList();

            // Switch to 1-based indices:
            row++;
            col++;

            // Throw exception if maximum coordinates will be out of bounds:
            if (OutOfBounds(row + (ulong)data.Count, col))
            {
                var message = String.Format("R{0}C{1} is out of bounds for this workbook type.", row + (ulong)data.Count, col);
                throw new Exceptions.NonFatals.GenericValueErrors(message, "XMLWorkbookGenerator::WriteColumn()");
            }
            if (data.Count > 0)
            {
                var writtenRows = this._ExistingSheetNamesToWrittenRows[this._ActiveSheet];
                foreach (var elem in data)
                {
                    // Throw exception if row has already been written to:
                    if (writtenRows.Contains(row))
                    {
                        throw new Exceptions.SemiFatals.FailedToCreateSheets(this._ActiveSheet, "XMLWorkbookGenerator::WriteColumn()", this._WorkbookName, String.Format("Row {0} has already been written to.", row));
                    }
                    // Update the row tracker:
                    this._ExistingSheetNamesToWrittenRows[this._ActiveSheet].Add(row);
                    var attributes = new List<OpenXmlAttribute>();
                    attributes.Add(new OpenXmlAttribute("r", null, row.ToString()));
                    // Write start tag for row:
                    this._HelperWriter.WriteStartElement(new Row(), attributes);
                    this.WriteToCell(row, col, elem);
                    // Write end tag for row:
                    this._HelperWriter.WriteEndElement();
                    row++;
                }
            }
        }
        /// <summary>
        /// Write to single cell on current selected sheet at passed coordinates.
        /// This method is kept private since does not write the necessary row tags.
        /// </summary>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <param name="data"></param>
        private void WriteToCell(ulong row, ulong col, string data, RangeFormatting? formats = null)
        {
            this._Attributes = new List<OpenXmlAttribute>();
            this._Attributes.Add(new OpenXmlAttribute("t", null, "str"));
            this._Attributes.Add(new OpenXmlAttribute("r", "", string.Format("{0}{1}", GetColumnName(col), row)));
            // Write cell start element with the type and reference attributes:
            this._HelperWriter.WriteStartElement(new Cell(), this._Attributes);
            // Write cell value:
            this._HelperWriter.WriteElement(new CellValue(data));
            // Finish writing cell:
            this._HelperWriter.WriteEndElement();
        }
        /// <summary>
        /// Write all data present in passed 2-dimensional container to current sheet, optionally starting at passed coordinates. 
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="AllData"></param>
        /// <param name="row"></param>
        /// <param name="col"></param>
        public void WriteAllData<T>(T AllData, ulong row = 0, ulong col = 0, RangeFormatting? formats = null) where T : IEnumerable<IEnumerable<string>>
        {
            // Write all data to sheet:
            foreach (var elem in AllData)
            {
                this.WriteRow(elem, row, col);
                row++;
            }
        }
        #endregion
        #region Interface Implementations
        /// <summary>
        /// Enable use with "using" clause. Will generate the file if GenerateFile was not called previously. 
        /// </summary>
        void IDisposable.Dispose()
        {
            this.GenerateFile();
        }
        #endregion
        #region Static Methods
        /// <summary>
        /// Indicate if the passed extension can be written to using XMLWriter object. 
        /// </summary>
        /// <param name="extension"></param>
        /// <returns></returns>
        public static bool ExtensionIsValid(string extension)
        {
            return XMLWorkbookGenerator._ValidExtensions.Contains(extension);
        }
        /// <summary>
        /// Convert Excel workbook column numeric index to alphabetical column name (ex: 1 = "A", 27 = "AA" etc). 
        /// columnIndex is expected to be 1 based.  
        /// </summary>
        /// <param name="columnIndex"></param>
        /// <returns></returns>
        private static string GetColumnName(ulong columnIndex)
        {
            int dividend = (int)columnIndex;
            string columnName = String.Empty;
            int modifier;
            while (dividend > 0)
            {
                modifier = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modifier).ToString() + columnName;
                dividend = (int)((dividend - modifier) / 26);
            }
            return columnName;
        }
        /// <summary>
        /// Determine if the passed address is within maximum row and column index. Depends upon what type of workbook is being generated.
        /// </summary>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <returns></returns>
        private bool OutOfBounds(ulong row, ulong col)
        {
            var maxDimensions = XMLWorkbookGenerator._MaxDimensions;
            return row >= maxDimensions.Item1 || col >= maxDimensions.Item2;
        }
        #endregion
    }
    /// <summary>
    /// Structure holds formatting information used for all cells in row or column.
    /// </summary>
    // See: https://stackoverflow.com/questions/39853019/how-to-make-some-text-bold-in-cell-using-openxml
    // https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2010/cc799272(v=office.14)
    [StructLayout(LayoutKind.Auto)]
    public struct RangeFormatting
    {
        #region Key Formatting Objects
        public CellFormat Format;
        public Run TextFormatting;
        #endregion
        #region Cell Attributes
        public uint? BorderID;
        public uint? CellFillColor;
        public HorizontalAlignmentValues? Alignment_Horizontal;
        public VerticalAlignmentValues? Alignment_Vertical;
        #endregion
        #region Font Attributes
        public bool? Bold;
        public bool? Italics;
        public uint? FontColor;
        public uint? FontSize;
        public uint? FontID;
        public uint? NumberFormat;
        #endregion
        #region Class Methods
        /// <summary>
        /// Generate all formatting objects to set fill and font formats in workbook.
        /// </summary>
        public void GenerateFormattingObjects()
        {
            this.Format = new CellFormat();
            this.TextFormatting = new Run();
            this.SetCellAttributes();
            this.SetFontAttributes();
        }
        /// <summary>
        /// Set the cell attributes for the range.
        /// </summary>
        private void SetCellAttributes()
        {
            // Set Cell Attributes:
            // Set Border ID if specified:
            if (this.BorderID.HasValue)
            {
                this.Format.BorderId = this.BorderID.Value;
                this.Format.ApplyBorder = true;
            }
            // Set fill color if specified:
            if (this.CellFillColor.HasValue)
            {
                this.Format.FillId = this.CellFillColor.Value;
                this.Format.ApplyFill = true;
            }
            // Set alignments if provided:
            if (this.Alignment_Horizontal.HasValue)
            {
                this.Format.Alignment.Horizontal = this.Alignment_Horizontal.Value;
                this.Format.ApplyAlignment = true;
            }
            if (this.Alignment_Vertical.HasValue)
            {
                this.Format.Alignment.Vertical = this.Alignment_Vertical.Value;
                this.Format.ApplyAlignment = true;
            }
        }
        /// <summary>
        /// Set the font attributes for the range.
        /// </summary>
        private void SetFontAttributes()
        {
            // Set Font Attributes:
            // Set the Run attribute if want to set font color, size and bold:
            if (this.FontColor.HasValue)
            {   

                this.Format.ApplyFont = true;
            }
            if (this.FontSize.HasValue)
            {
                

            }
            if (this.Bold.HasValue && this.Bold.Value)
            {
                this.TextFormatting.RunProperties.Append(new Bold());
            }

            // Set the CellFormat attributes using all remaining properties:
            if (this.FontID.HasValue)
            {
                this.Format.FontId = this.FontID.Value;
                this.Format.ApplyFont = true;
            }
            
            if (this.NumberFormat.HasValue)
            {
                this.Format.NumberFormatId = this.NumberFormat.Value;
                this.Format.ApplyNumberFormat = true;
            }
        }
        #endregion
    }
    #endregion
}
