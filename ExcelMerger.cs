using System;
using System.Collections.Generic;
using System.Data;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelMerge {
    public class ExcelMerger {
        public string Filename { get; set; }
        private HashSet<string> _files;
        private Excel._Application _app;
        private Excel._Workbook _workbook;
        private Excel._Worksheet _worksheet;


        public ExcelMerger() {
            Filename = $"untitled - {DateTime.Today}";
            _files = new HashSet<string>();
            _app = new Excel.Application();
            _workbook = _app.Workbooks.Add();
            _worksheet = (Excel.Worksheet) _workbook.ActiveSheet;
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

            /* Match the contents of all subsequent files to the headers pulled from the first file */
        }

        public void _readWorkbook(string filename, bool firstWorkbook = false) {
            int row = (firstWorkbook) ? 1 : 2;
            int col = 1;

            Excel._Workbook tempWorkbook = _app.Workbooks.Open(filename, ReadOnly: true);
            Excel._Worksheet tempWorksheet = (Excel._Worksheet) tempWorkbook.ActiveSheet;
            Excel.Range usedRange = tempWorksheet.UsedRange;

            for (int i = row; i <= usedRange.Rows.Count; i++) {
                for (int j = col; j <= usedRange.Columns.Count; j++) {
                    Console.Write(((Excel.Range)usedRange.Item[i, j]).Value.ToString());
                }
                Console.WriteLine();
            }
            tempWorkbook.Close();
        }

        public void Export() {
            _workbook.SaveAs(Filename);
            
        }

        public void ExportAs(string filename) {
            Filename = filename;
            Export();
        }

        public void Quit() {
            _app.Quit();
        }

        public static DataTable WorksheetToDataTable(string filename) {
            
            return null;
        }

        private void _readFirstRow() { }
    }
}