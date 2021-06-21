using System;
using ExcelMerge;
using NUnit.Framework;

namespace EMTest {
    [TestFixture]
    public class ExcelMergerTest {
        private const string INPUT_FOLDER = 
            @"C:\Users\Basestar\RiderProjects\ExcelMerge\EMTest\TestResources\TR_SOURCE\";

        private const string OUTPUT_FOLDER =
            @"C:\Users\Basestar\RiderProjects\ExcelMerge\EMTest\TestResources\TR_RESULT\";
        
        private static string[] _testFiles = {
            INPUT_FOLDER + @"TR_1.xlsx",
            INPUT_FOLDER + @"TR_2.xlsx"
        };

        private static string[] _testFiles2 = {
            INPUT_FOLDER + @"20124B-A500.xls",
            INPUT_FOLDER + @"20124B-A501.xls",
            INPUT_FOLDER + @"20124B-A502.xls",
            INPUT_FOLDER + @"20124B-A503.xls",
            INPUT_FOLDER + @"20124B-A504.xls",
            INPUT_FOLDER + @"20124B-A505.xls"
        };

        [Test]
        public void checkTR() {
            foreach (var VARIABLE in _testFiles) {
                Console.WriteLine(VARIABLE);
            }
        }

        [Test]
        public void AddWorkbook() {
            ExcelMerger em = new ExcelMerger();
            em.AddWorkbook(INPUT_FOLDER + _testFiles[0]);
            Assert.True(em.Files.Count == 1);
            em.Quit();
            
            
        }

        [Test]
        public void AddWorkbooks() {
            ExcelMerger em = new ExcelMerger();
            em.AddWorkbooks(_testFiles);
            Assert.True(em.Files.Count == 2);
            em.Quit();
        }

        [Test]
        public void Merge() {
            ExcelMerger em = new ExcelMerger();
            em.AddWorkbooks(_testFiles);
            em.Merge();
            Console.Write(em.NewWorksheet.Count);
            Assert.True(em.NewWorksheet.Count == 7);
            em.Quit();
        }

        [Test]
        public void Slim() {
            ExcelMerger em = new ExcelMerger();
            em.AddWorkbooks(_testFiles);
            em.Merge();
            em.Slim(0, summableFields: new int[]{1});
            Console.Write(em.NewWorksheet.Count);
            Assert.True(em.NewWorksheet.Count == 6);
            em.Quit();
        }

        [Test]
        public void Export() {
            ExcelMerger em = new ExcelMerger();
            em.AddWorkbooks(_testFiles);
            em.Filename = OUTPUT_FOLDER + "RESULT.xlsx";
            em.Merge();
            em.Slim(0, summableFields: new int[]{1});
            em.Export();
            Console.Write(em.NewWorksheet.Count);
            Assert.True(em.NewWorksheet.Count == 6);
            em.Quit();
        }

        [Test]
        public void ExportAs() {
            ExcelMerger em = new ExcelMerger();
            em.AddWorkbooks(_testFiles);
            em.Merge();
            em.Slim(0, summableFields: new int[]{1});
            //em.ExportAs("RESULT");
            Console.Write(em.NewWorksheet.Count);
            Assert.True(em.NewWorksheet.Count == 6);
            em.Quit();
        }

        [Test]
        public void ExportAs2() {
            ExcelMerger em = new ExcelMerger();
            em.AddWorkbooks(_testFiles2);
            em.Merge();
            em.Slim(2, 6, new int[]{1});
            em.ExportAs(OUTPUT_FOLDER + "RESULT2.xlsx");
            em.Quit();
        }

        [Test]
        public void Quit() {
            
        }
    }
}