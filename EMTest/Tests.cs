using System;
using System.Collections.Generic;
using ExcelMerge;
using NUnit.Framework;
using Excel = Microsoft.Office.Interop.Excel;

namespace EMTest {
    [TestFixture]
    public class Tests {
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
            em.Filename = "TEST3";
            em._readWorkbook(@"C:\Users\Basestar\Documents\TEST.xlsx", true);
            em.Export();
            em.Quit();
        }

        [Test]
        public void ReadWorkbookMultiBook() {
            ExcelMerger em = new ExcelMerger();
            em.Filename = "TEST4";
            em._readWorkbook(@"C:\Users\Basestar\Documents\TEST.xlsx", true);
            em._readWorkbook(@"C:\Users\Basestar\Documents\TEST.xlsx");
            em.Export();
            em.Quit();
        }

        [Test]
        public void Merge() {
            ExcelMerger em = new ExcelMerger();
            em.Filename = "TEST5";
            em.AddWorkbooks(new []{@"C:\Users\Basestar\Documents\TEST.xlsx", @"C:\Users\Basestar\Documents\TEST6.xlsx"});
            em.Merge();
            em.Export();
            em.Quit();
        }

        [Test]
        public void PrintNewWorkbook() {
            ExcelMerger em = new ExcelMerger();
            em.Filename = "TEST5";
            em.AddWorkbook(@"C:\Users\Basestar\Documents\TEST6.xlsx");
            em.Merge();


            // for (int i = 0; i < em.NewWorksheet.Count; i++) {
            //     SortedList<int, object> val = null;
            //     
            //     if (em.NewWorksheet.TryGetValue(i, out val)) {
            //         for (int j = 0; j < val.Count; j++) {
            //             object output = null;
            //             if (val.TryGetValue(j, out output)) {
            //                 Console.Write($"| {output}");
            //             }
            //             
            //         }
            //         Console.WriteLine();
            //     }
            // }
            foreach (var lis in em.NewWorksheet) {
                foreach (var obj in lis) {
                    Console.Write($"| {obj} ");
                }
                Console.WriteLine();
            }

            em.Quit();
            
        }

        [Test]
        public void SlimNewWorkbook() {
            ExcelMerger em = new ExcelMerger();
            em.Filename = "TEST7";
            em.AddWorkbook(@"C:\Users\Basestar\Documents\TEST6.xlsx");
            em.Merge();
            em.Slim(0, summableFields: new []{1});

            foreach (var lis in em.NewWorksheet) {
                foreach (var obj in lis) {
                    Console.Write($"| {obj} ");
                }
                Console.WriteLine();
            }
            em.Export();
            em.Quit();
        }

        [Test]
        public void FullMergeSlimTest() {
            ExcelMerger em = new ExcelMerger();
            em.Filename = "FULL_TEST_1";
            em.AddWorkbooks(new [] {
                @"C:\Users\Basestar\Documents\20124B-A505.xls",
                @"C:\Users\Basestar\Documents\20124B-A503.xls"
            });
            em.Merge();
            em.Slim(3, summableFields:new []{1});
            foreach (var lis in em.NewWorksheet) {
                foreach (var obj in lis) {
                    Console.Write($"| {obj} ");
                }
                Console.WriteLine();
            }
            em.Export();
            em.Quit();
        }
        [Test]
        public void FullMergeSlimTest2() {
            ExcelMerger em = new ExcelMerger();
            em.Filename = "FULL_TEST_3";
            em.AddWorkbooks(new [] {
                @"C:\Users\Basestar\Documents\20124B-A505.xls",
                @"C:\Users\Basestar\Documents\20124B-A504.xls"
            });
            em.Merge();
            em.Slim(2, 6, new []{1});
            foreach (var lis in em.NewWorksheet) {
                foreach (var obj in lis) {
                    Console.Write($"| {obj} ");
                }
                Console.WriteLine();
            }
            em.Export();
            em.Quit();
        }
    }
}