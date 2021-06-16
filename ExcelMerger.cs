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
            var tempWorkbook = _app.Workbooks.Open(enumerator.Current, ReadOnly: true);
            var tempWorksheet = (Excel._Worksheet) tempWorkbook.ActiveSheet;
            

            /* Read the first row headers of the first file */

            /* Match the contents of all subsequent files to the headers pulled from the first file */
        }

        private string _itemSelect(int row, int column) {
            
            return null;
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

        private void _readFirstRow() { }
    }
}