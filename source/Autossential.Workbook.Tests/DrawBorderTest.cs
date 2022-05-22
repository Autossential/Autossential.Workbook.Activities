using Autossential.Shared.Tests;
using Autossential.Workbook.Core;
using Autossential.Workbook.Core.Enums;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Drawing;

namespace Autossential.Workbook.Tests
{
    [TestClass]
    public class DrawBorderTest
    {
        [ClassInitialize]
        public static void Initialize(TestContext _)
        {
            IOSamples.ClearFolder();

            IOSamples.ExportSample("empty.xlsx", "borders.xlsx");
       //     IOSamples.ExportSample("empty.xls", "borders.xls");
        }

        [ClassCleanup]
        public static void Clean()
        {
            IOSamples.CopyToDownloadsFolder(IOSamples.GetTestPath("borders.xlsx"));
      //      IOSamples.CopyToDownloadsFolder(IOSamples.GetTestPath("borders.xls"));
        }

        [TestMethod]
        //[DataRow(Border.All, "borders.xlsx")]
        //[DataRow(Border.Top, "borders.xlsx")]
        //[DataRow(Border.Bottom, "borders.xlsx")]
        //[DataRow(Border.Left, "borders.xlsx")]
        //[DataRow(Border.Right, "borders.xlsx")]
        [DataRow(Border.Outside, "borders.xlsx")]
        // [DataRow("borders.xls")]
        public void OpenXml(Border border, string fileName)
        {
            var path = IOSamples.GetTestPath(fileName);
            var adapter = WorkbookAdapterFactory.Create(path);
            adapter.DrawBorder(border.ToString(), "E5:I15", border, BorderStyle.Thick, Color.Blue);
            adapter.Save();
            adapter.Dispose();
        }
    }
}
