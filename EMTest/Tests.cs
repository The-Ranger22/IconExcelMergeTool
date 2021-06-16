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
            var obj = (string)(ws.Cells[1, 1] as Excel.Range).Value;
            wb.Close();
            wb = excelApp.Workbooks.Add();
            Excel._Worksheet workSheet = (Excel.Worksheet) wb.ActiveSheet;
            workSheet.Cells[1, "A"] = obj;
            
            wb.SaveAs("TEST2");
            wb.Close();
            excelApp.Quit();
            
        }
        
    }
}