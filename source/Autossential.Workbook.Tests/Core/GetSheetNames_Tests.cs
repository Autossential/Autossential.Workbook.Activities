using Autossential.Shared.Tests;
using Autossential.Workbook.Core;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Data;

namespace Autossential.Workbook.Tests.Core
{
    [TestClass]
    public class GetSheetNames_Tests
    {

        [TestMethod]
        [DataRow("OXML_sheets.xlsx")]
        [DataRow("BIFF8_sheets.xls")]
        public void GetSheetNames_MultipleSheets_ReturnSheetNames(string fileName)
        {
            var path = IOSamples.GetSamplePath(fileName);
            var workbook = WorkbookProcessorFactory.OpenOrCreate(path);
            var sheetNames = workbook.GetSheetNames();
            workbook.Dispose();
            CollectionAssert.AreEquivalent(sheetNames, new[] { "Sheet1", "Sheet2", "Sheet3", "Sheet4" });
        }
    }
}
