using Autossential.Shared.Tests;
using Autossential.Workbook.Core;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Data;
using System.Linq;

namespace Autossential.Workbook.Tests.Core
{
    [TestClass]
    public class ReadRange_Tests
    {
        [TestMethod]
        [DataRow("OXML_data.xlsx", "A1", true, 10, 6)]
        [DataRow("OXML_data.xlsx", "A1", false, 11, 6)]
        [DataRow("OXML_data.xlsx", "A1:B5", true, 4, 2)]

        [DataRow("BIFF8_data.xls", "A1", true, 10, 6)]
        [DataRow("BIFF8_data.xls", "A1", false, 11, 6)]
        [DataRow("BIFF8_data.xls", "A1:B5", true, 4, 2)]
        public void ReadRange_SheetNameExists_ReturnRange(string fileName, string range, bool hasHeaders, int expectedRowCount, int expectedColumnCount)
        {
            var path = IOSamples.GetSamplePath(fileName);
            var workbook = WorkbookProcessorFactory.OpenOrCreate(path);
            var dataTable = workbook.ReadRange("Sheet1", range, hasHeaders, false);
            workbook.Dispose();
            Assert.AreEqual(expectedRowCount, dataTable.Rows.Count);
            Assert.AreEqual(expectedColumnCount, dataTable.Columns.Count);
        }

        [TestMethod]
        [DataRow("OXML_sheets.xlsx")]
        [DataRow("BIFF8_sheets.xls")]
        public void ReadRange_SheetNameDoesNotExist_ThrowException(string fileName)
        {
            var path = IOSamples.GetSamplePath(fileName);
            var workbook = WorkbookProcessorFactory.OpenOrCreate(path);
            Assert.ThrowsException<ArgumentException>(() => workbook.ReadRange("DoesNotExist", "A1", false, false));
            workbook.Dispose();
        }

        [TestMethod]
        [DataRow("OXML_sheets.xlsx")]
        [DataRow("BIFF8_sheets.xls")]
        public void ReadRange_InvalidSheetName_ThrowsArgumentException(string fileName)
        {
            var path = IOSamples.GetSamplePath(fileName);
            var workbook = WorkbookProcessorFactory.OpenOrCreate(path);
            Assert.ThrowsException<ArgumentException>(() => workbook.ReadRange("", "A1", false, false));
            workbook.Dispose();
        }

        [TestMethod]
        [DataRow("OXML_sheets.xlsx")]
        [DataRow("BIFF8_sheets.xls")]
        public void ReadRange_InvalidRange_ThrowsArgumentException(string fileName)
        {
            var path = IOSamples.GetSamplePath(fileName);
            var workbook = WorkbookProcessorFactory.OpenOrCreate(path);
            Assert.ThrowsException<ArgumentException>(() => workbook.ReadRange("Sheet1", "A", false, false));
            workbook.Dispose();
        }

        [TestMethod]
        [DataRow("OXML_data.xlsx", false)]
        [DataRow("OXML_data.xlsx", true)]
        public void ReadRange_ColumnTypes_ReturnRange(string fileName, bool useColumnTypes)
        {
            var path = IOSamples.GetSamplePath(fileName);
            var workbook = WorkbookProcessorFactory.OpenOrCreate(path);
            var dataTable = workbook.ReadRange("Sheet1", "A1", true, useColumnTypes);
            workbook.Dispose();

            Assert.AreEqual(6, dataTable.Columns.Count);
            var columnTypes = dataTable.Columns.Cast<DataColumn>().Select(p => p.DataType).ToArray();

            if (useColumnTypes)
            {
                CollectionAssert.AreEqual(columnTypes, new[]
                {
                    typeof(int),
                    typeof(decimal),
                    typeof(decimal),
                    typeof(DateTime),
                    typeof(string),
                    typeof(bool)
                });
            }
            else
            {
                Assert.IsTrue(columnTypes.All(p => p == typeof(object)));
            }
        }
    }
}
