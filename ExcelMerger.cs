using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelMerge {
    /// <summary>
    /// ExcelMerger is a lightweight class designed to manipulate excel spreadsheets, primarily merging them.
    /// </summary>
    /// <remarks>
    /// For the sake of expediency, the ExcelMerger class relies as little as possible on excel when manipulating data, only interacting with it during file read/write.
    /// </remarks>
    public class ExcelMerger {
        public const int NULL_KEY = -1;
        public string Filename { get; set; }
        public HashSet<ArrayList> NewWorksheet { get; set; }
        public ArrayList Header { get; set; }
        public HashSet<string> Files { get; }
        private Excel._Application _app;
        private Excel._Workbook _workbook;
        private Excel._Worksheet _worksheet;
        private bool _closed;
        public ArrayList Ignorables { get; private set; }

        public ExcelMerger() {
            Filename = $"untitled - {DateTime.Today}";
            Files = new HashSet<string>();
            _app = new Excel.Application();
            _workbook = _app.Workbooks.Add();
            _worksheet = (Excel.Worksheet) _workbook.ActiveSheet;
            _closed = false;
            NewWorksheet = new HashSet<ArrayList>();
            Ignorables = new ArrayList();
        }

        public void AddWorkbook(string filename) {
            Files.Add(filename);
        }

        public void AddWorkbooks(string[] filenames) {
            foreach (string filename in filenames) {
                AddWorkbook(filename);
            }
        }

        public void Merge() {
            var enumerator = Files.GetEnumerator();

            /* Read the first row headers of the first file */
            enumerator.MoveNext(); // Move to point at the first item in the hash set.
            _readWorkbook(enumerator.Current, true); // read the contents of the first workbook
            /* Match the contents of all subsequent files to the headers pulled from the first file */
            while (enumerator.MoveNext()) {
                _readWorkbook(enumerator.Current);
            }
            
            enumerator.Dispose();
        }

        public void Slim(int primaryKey, int secondaryKey = NULL_KEY, int[] summableFields = null) {
            HashSet<ArrayList> slimmedWorksheet = new HashSet<ArrayList>();
            foreach (var lis in NewWorksheet) {
                bool preexistingKey = false;
                //if slimmedWorksheet does not already contain a list containing the designated primary key, add the list to slimmedWorksheet
                if (!slimmedWorksheet.Contains(lis)) {
                    foreach (var record in slimmedWorksheet) {
                        //check if the primary key 
                        if (!(record[primaryKey] is null) && !Ignorables.Contains(record[primaryKey].ToString()) && record[primaryKey].Equals(lis[primaryKey])) {
                            foreach (var index in summableFields) {
                                record[index] = (double) (lis[index]) + (double) (record[index]);
                            }

                            preexistingKey = true;
                            break;
                        }
                        /*
                         * If the record at the given primary key is null AND
                         * the secondary key is not null AND
                         * the record at the given secondary key is not null AND
                         * the value at the record at the given secondary key is not an ignorable AND
                         * the values given by the secondary key in both the record and the list are equal,
                         * sum
                         */
                        if (
                            (record[primaryKey] is null ||  Ignorables.Contains(record[primaryKey].ToString())) && 
                            secondaryKey != NULL_KEY &&
                            !(record[secondaryKey] is null) && 
                            !Ignorables.Contains(record[secondaryKey].ToString()) &&
                            record[secondaryKey].Equals(lis[secondaryKey])
                            ) {
                            foreach (var index in summableFields) {
                                record[index] = (double) (lis[index]) + (double) (record[index]);
                            }
                            preexistingKey = true;
                            break;
                        }
                    }

                    if (!preexistingKey) {
                        slimmedWorksheet.Add(lis);
                    }
                }
            }

            NewWorksheet = slimmedWorksheet;
        }
        
        /// <summary>
        /// Reads the first worksheet from an excel workbook into a new workbook.
        /// </summary>
        /// <param name="filename">The filename of the excel spreadsheet that is to be read.</param>
        /// <param name="firstWorkbook">If true, the first row of the worksheet will be included in the read, setting the headers of the new workbook. If false, the first row of the workbook is ignored.</param>
        public void _readWorkbook(string filename, bool firstWorkbook = false) {
            int row = (firstWorkbook) ? 1 : 2; // The starting row
            int col = 1; // The starting column
            
            try {
                Excel._Workbook tempWorkbook = _app.Workbooks.Open(filename, ReadOnly: true);
                Excel._Worksheet tempWorksheet = (Excel._Worksheet) tempWorkbook.ActiveSheet;
                Excel.Range usedRange = tempWorksheet.UsedRange;
                //Excel.Range newBookRange = _worksheet.Cells;

                for (int i = row; i <= usedRange.Rows.Count; i++) {
                    ArrayList lis = new ArrayList();
                    for (int j = col; j <= usedRange.Columns.Count; j++) {
                        //Console.Write(((Excel.Range) usedRange.Item[i, j]).Value.ToString());
                        lis.Add(((Excel.Range) usedRange.Item[i, j]).Value);
                        //newBookRange.Item[_depth, j] = ((Excel.Range) usedRange.Item[i, j]).Value;
                    }

                    if (i == 1) {
                        Header = lis;
                    }
                    NewWorksheet.Add(lis);
                }
                
                tempWorkbook.Close();
            }
            catch (COMException e) {
                Console.Write($"Could not find file! {e}");
            }
        }

        public void ParseIgnorables(string str, char delimiter) {
            Ignorables.AddRange(str.Split(delimiter));
        }

        public void Export() {
            int rowCounter = 1;
            foreach (var lis in NewWorksheet) {
                int colCounter = 1;
                foreach (var obj in lis) {
                    _worksheet.Cells.Item[rowCounter, colCounter] = obj;
                    _worksheet.Cells.ColumnWidth = 20;
                    colCounter++;
                }
                rowCounter++;
            }
            
            _workbook.SaveAs(Filename);
            _workbook.Close();
            _closed = true;
        }

        /// <summary>
        /// Save a new excel file under the specified file name.
        /// </summary>
        /// <param name="filename"></param>
        public void ExportAs(string filename) {
            Filename = filename;
            Export();
        }

        public void Quit() {
            if (!_closed) {
                _workbook.Close(false);
            }
            _app.Quit();
            
        }
    }
}