using System;
using ExcelMerge;
using NUnit.Framework;
using Excel = Microsoft.Office.Interop.Excel;

namespace EMTest {
    [TestFixture]
    public class Tests {
        ExcelMerger _em = new ExcelMerger();

        [Test]
        public void Test1() {
            var excelApp = new Excel.Application();
            excelApp.Visible = true;
            Excel._Workbook wb = excelApp.Workbooks.Add();
            Excel._Worksheet workSheet = (Excel.Worksheet) wb.ActiveSheet;
            workSheet.Cells[1, "A"] = "ID Number";
            workSheet.Cells[1, "B"] = "Current Balance";
            var row = 1;
            Tuple<string, int>[] accounts = new[] {
                new Tuple<string, int>("ALPHA", 1),
                new Tuple<string, int>("BETA", 2),
                new Tuple<string, int>("GAMMA", 3)
            };
            foreach (var acct in accounts) {
                row++;
                workSheet.Cells[row, "A"] = acct.Item1;
                workSheet.Cells[row, "B"] = acct.Item2;
            }

            wb.SaveAs("TEST");
            wb.Close();
            excelApp.Quit();
            Assert.True(true);
        }

        [Test]
        public void Test2() {
            var excelApp = new Excel.Application();
            excelApp.Visible = false;
            Excel._Workbook wb = excelApp.Workbooks.Open("C:\\Users\\Basestar\\Documents\\TEST.xlsx");
            Excel._Worksheet ws = (Excel._Worksheet) wb.ActiveSheet;
            Excel.Range usedRange = ws.UsedRange;
            Console.Write($"Rows : {usedRange.Rows.Count} | Columns : {usedRange.Columns.Count}");
            // int currentCol = 1;
            // var val = (string) ((Excel.Range) cols[1, currentCol]).Value;
            //
            // while (val != null) {
            //     Console.Write(val);
            //     
            //     val = (string) ((Excel.Range) cols[1, ++currentCol]).Value;
            // }


            wb.Close();
            excelApp.Quit();
            
        }

        [Test]
        public void ReadWorkbookTest() {
            ExcelMerger em = new ExcelMerger();
            em._readWorkbook(@"C:\Users\Basestar\Documents\TEST.xlsx");
            em.Quit();
        }

    }
}