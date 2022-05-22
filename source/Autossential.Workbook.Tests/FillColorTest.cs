using Autossential.Shared.Tests;
using Autossential.Workbook.Core;
using Autossential.Workbook.Core.Enums;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Drawing;

namespace Autossential.Workbook.Tests
{
    [TestClass]
    public class FillColorTest
    {
        [ClassInitialize]
        public static void Initialize(TestContext _)
        {
            IOSamples.ClearFolder();

            var xlsx = WorkbookAdapterFactory.Create(IOSamples.ExportSample("empty.xlsx", "colors.xlsx"));
            xlsx.RenameSheet(0, "Horizontal");
            xlsx.Save();
            xlsx.Dispose();

            var xls = WorkbookAdapterFactory.Create(IOSamples.ExportSample("empty.xls", "colors.xls"));
            xls.RenameSheet(0, "Horizontal");
            xls.Save();
            xls.Dispose();
        }

        [ClassCleanup]
        public static void Clean()
        {
            IOSamples.CopyToDownloadsFolder(IOSamples.GetTestPath("colors.xlsx"));
            IOSamples.CopyToDownloadsFolder(IOSamples.GetTestPath("colors.xls"));
        }

        [TestMethod]
        [DataRow("colors.xlsx")]
        [DataRow("colors.xls")]
        public void HorizontalFill(string fileName)
        {
            var path = IOSamples.GetTestPath(fileName);
            var adapter = WorkbookAdapterFactory.Create(path);
            adapter.FillColor("Horizontal", "A1:E5", new[] { Color.Teal, Color.Orange }, FillOrientation.Horizontal);
            adapter.Save();
            adapter.Dispose();
        }

        [TestMethod]
        [DataRow("colors.xlsx")]
        [DataRow("colors.xls")]
        public void VerticalFill(string fileName)
        {
            var path = IOSamples.GetTestPath(fileName);
            var adapter = WorkbookAdapterFactory.Create(path);
            adapter.FillColor("Vertical", "A1:E5", new[] { Color.Teal, Color.Orange }, FillOrientation.Vertical);
            adapter.Save();
            adapter.Dispose();
        }

        [TestMethod]
        [DataRow("colors.xlsx")]
        [DataRow("colors.xls")]
        public void ChessFill(string fileName)
        {
            var path = IOSamples.GetTestPath(fileName);
            var adapter = WorkbookAdapterFactory.Create(path);
            adapter.FillColor("Chess", "A1:E5", new[] { Color.Teal, Color.Orange }, FillOrientation.Chess);
            adapter.Save();
            adapter.Dispose();
        }

        [TestMethod]
        [DataRow("colors.xlsx")]
        [DataRow("colors.xls")]
        public void DiagonalLeft(string fileName)
        {
            var path = IOSamples.GetTestPath(fileName);
            var adapter = WorkbookAdapterFactory.Create(path);
            adapter.FillColor("DiagonalLeft", "A1:E5", new[] { Color.Teal, Color.Orange }, FillOrientation.DiagonalLeft);
            adapter.Save();
            adapter.Dispose();
        }

        [TestMethod]
        [DataRow("colors.xlsx")]
        [DataRow("colors.xls")]
        public void DiagonalRight(string fileName)
        {
            var path = IOSamples.GetTestPath(fileName);
            var adapter = WorkbookAdapterFactory.Create(path);
            adapter.FillColor("DiagonalRight", "A1:E5", new[] { Color.Teal, Color.Orange }, FillOrientation.DiagonalRight);
            adapter.Save();
            adapter.Dispose();
        }
    }
}