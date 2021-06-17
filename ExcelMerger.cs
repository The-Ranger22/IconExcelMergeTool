using System;
using System.Collections.Generic;
using System.Data;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelMerge {
    public class ExcelMerger {
        public string Filename { get; set; }
        public SortedDictionary<int, SortedList<int, object>> NewWorksheet { get; set; }
        private HashSet<string> _files;
        private Excel._Application _app;
        private Excel._Workbook _workbook;
        private Excel._Worksheet _worksheet;
        private int _depth;
        private bool _closed;
        


        public ExcelMerger() {
            Filename = $"untitled - {DateTime.Today}";
            _files = new HashSet<string>();
            _app = new Excel.Application();
            _workbook = _app.Workbooks.Add();
            _worksheet = (Excel.Worksheet) _workbook.ActiveSheet;
            _depth = 1;
            _closed = false;
            NewWorksheet = new SortedDictionary<int, SortedList<int, object>>();
        }

        public void AddWorkbook(string filename) {
            _files.Add(filename);
        }

        public void AddWorkbooks(string[] filenames) {
            foreach (string filename in filenames) {
                AddWorkbook(filename);
            }
        }

        public void Merge() {
            var enumerator = _files.GetEnumerator();

            /* Read the first row headers of the first file */
            enumerator.MoveNext(); // Move to point at the first item in the hash set.
            _readWorkbook(enumerator.Current, true); // read the contents of the first workbook
            /* Match the contents of all subsequent files to the headers pulled from the first file */
            while (enumerator.MoveNext()) {
                _readWorkbook(enumerator.Current);
            }
            enumerator.Dispose();
        }

        public void Slim(int primaryKey) {
            SortedDictionary<int, SortedList<int, object>> slimmedWorksheet =
                new SortedDictionary<int, SortedList<int, object>>();
            foreach (var lis in NewWorksheet.Values) {
                //if slimmedWorksheet does not already contain a list containing the designated primary key, add the list to slimmedWorksheet
                if (slimmedWorksheet.ContainsValue(lis)) {
                    
                }
                //else, update the integer values already stored in the record owned by the primary key
                
                foreach (var obj in lis.Values) {
                    Console.Write($"| {obj} ");
                }
                Console.WriteLine();
            }
            
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
                Excel.Range newBookRange = _worksheet.Cells;

                for (int i = row; i <= usedRange.Rows.Count; i++) {
                    SortedList<int, object> lis = new SortedList<int, object>();
                    for (int j = col; j <= usedRange.Columns.Count; j++) {
                        //Console.Write(((Excel.Range) usedRange.Item[i, j]).Value.ToString());
                        lis.Add(j, ((Excel.Range) usedRange.Item[i, j]).Value);
                        newBookRange.Item[_depth, j] = ((Excel.Range) usedRange.Item[i, j]).Value;
                    }
                    NewWorksheet.Add(_depth, lis);
                    _depth++;
                    Console.WriteLine();
                }

                tempWorkbook.Close();
            }
            catch (COMException e) {
                Console.Write($"Could not find file! {e}");
            }
        }

        public void Export() {
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

        public static DataTable WorksheetToDataTable(string filename) {
            return null;
        }

        private void _readFirstRow() { }
        
        
        
        
        
    }
}