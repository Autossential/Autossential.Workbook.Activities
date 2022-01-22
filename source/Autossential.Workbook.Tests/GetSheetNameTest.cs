using Autossential.Shared.Tests;
using Autossential.Workbook.Activities;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;

namespace Autossential.Workbook.Tests
{
    [TestClass]
    public class GetSheetNameTest
    {
        [ClassInitialize]
        public static void Initialize(TestContext _)
        {
            IOSamples.ClearFolder();
            IOSamples.ExportSample("book.xls");
            IOSamples.ExportSample("book.xlsx");
        }

        [ClassCleanup]
        public static void CleanUp()
        {
            IOSamples.ClearFolder();
        }

        [TestMethod]
        [DataRow("book.xlsx", 0, "Sheet1")]
        [DataRow("book.xlsx", 1, "Sheet2")]
        public void OpenXML(string file, int sheetIndex, string expectedSheetName)
        {
            var path = IOSamples.GetTestPath(file);
            var args = new Dictionary<string, object>
            {
                { nameof(GetSheetName.WorkbookPath), path },
                { nameof(GetSheetName.SheetIndex), sheetIndex },
            };

            var workflow = new GetSheetName
            {
                UseScope = false
            };

            var result = WorkflowTester.Invoke(workflow, args);
            Assert.AreEqual(expectedSheetName, result);
        }

        [TestMethod]
        [DataRow("book.xls", 2, "Sheet3")]
        [DataRow("book.xls", 3, "Sheet4")]
        public void OLE2(string file, int sheetIndex, string expectedSheetName)
        {
            var path = IOSamples.GetTestPath(file);
            var args = new Dictionary<string, object>
            {
                { nameof(GetSheetName.WorkbookPath), path },
                { nameof(GetSheetName.SheetIndex), sheetIndex },
            };

            var workflow = new GetSheetName
            {
                UseScope = false
            };

            var result = WorkflowTester.Invoke(workflow, args);
            Assert.AreEqual(expectedSheetName, result);
        }
    }
}
