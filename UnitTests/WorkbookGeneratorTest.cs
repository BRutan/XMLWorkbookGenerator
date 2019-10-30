/* WorkbookGeneratorTest.cs
 Description:
    * Test the XMLWorkbookGenerator wrapper class by writing large number of elements to multiple sheets of test workbook.
*/

using Microsoft.VisualBasic;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Diagnostics;
using System.Threading.Tasks;
using CSharpObjectLibrary.OpenXMLTools;
using LibraryExceptions = CSharpObjectLibrary.Exceptions;

namespace UnitTests
{
    class WorkbookGeneratorTest
    {
        static void Main(string[] args)
        {
            // Get the relevant file path to output the test workbook:
            string filePath = Interaction.InputBox("Please enter path to output test workbook (include file name):", "XMLWorkbookGenerator Test:");
            uint numRows = 10000, numCols = 1000;

            // Display exception if no file path was provided:
            if (String.IsNullOrWhiteSpace(filePath))
            {
                System.Windows.Forms.MessageBox.Show("No output folder provided. Exiting test program.");
                System.Environment.Exit(1);
            }

            // Generate the workbook, calculate completion time:
            var sheetNames = new List<string>() { "First", "Second", "Third" };
            var stopwatch = new Stopwatch();
            try
            {
                using (XMLWorkbookGenerator workbook = new XMLWorkbookGenerator(filePath))
                {
                    string[] Content = new string[numCols];

                    for (int col = 0; col < numCols; col++)
                    {
                        Content[col] = col.ToString();
                    }
                    stopwatch.Start();
                    foreach (var sheetName in sheetNames)
                    {
                        workbook.AddSheet(sheetName);
                        // Write data to current sheet:
                        for (uint rowNum = 0; rowNum < numRows; rowNum++)
                        {
                            workbook.WriteRow(Content, rowNum);
                        }
                        workbook.FinishSheet();
                    }
                }
            }
            catch(LibraryExceptions.ExceptionAggregator agg)
            {
                // Display issues if any exceptions occurred, then exit script:
                System.Windows.Forms.MessageBox.Show(agg.ConciseMessage());
                System.Environment.Exit(1);
            }
            stopwatch.Stop();
            string message = String.Format("Completion time for {0} rows, {1} columns over {2} sheets: {3}", numRows, numCols, sheetNames.Count, stopwatch.Elapsed.ToString("mm\\:ss\\.ff"));
            System.Windows.Forms.MessageBox.Show(message);
        }
    }
}
