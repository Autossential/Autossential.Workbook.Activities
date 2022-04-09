using Autossential.Shared.Tests;
using Autossential.Workbook.Core;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.IO;

namespace Autossential.Workbook.Tests
{
    [TestClass]
    public class FreezeUnfreezePanesTest
    {
        [TestMethod]
        [DataRow(1, 0)]
        [DataRow(0, 1)]
        [DataRow(2, 2)]
        public void OpenXML_FreezeUnfreeze(int cols, int rows)
        {
            var filePath = CopyToDownloadsFolder(IOSamples.GetSamplePath("openxml.xlsx"), $"c{cols}r{rows}_freeze.xlsx");
            var adapter = WorkbookAdapterFactory.Create(filePath);
            adapter.FreezePanes("Sheet1", cols, rows);
            adapter.Save();
            adapter.Dispose();

            Unfreeze(filePath);
        }

        private void Unfreeze(string filePath)
        {
            var unfreeze = filePath.Replace("_freeze", "_unfreeze");
            File.Copy(filePath, unfreeze, true);
            var adapter = WorkbookAdapterFactory.Create(unfreeze);
            adapter.FreezePanes("Sheet1", 0, 0);
            adapter.Save();
            adapter.Dispose();
        }

        [TestMethod]
        [DataRow(1, 0)]
        [DataRow(0, 1)]
        [DataRow(2, 2)]
        public void OLE2_FreezeUnfreeze(int cols, int rows)
        {
            var filePath = CopyToDownloadsFolder(IOSamples.GetSamplePath("ole2.xls"), $"c{cols}r{rows}_freeze.xls");
            var adapter = WorkbookAdapterFactory.Create(filePath);
            adapter.FreezePanes("Sheet1", cols, rows);
            adapter.Save();
            adapter.Dispose();

            Unfreeze(filePath);
        }

        private static string CopyToDownloadsFolder(string filePath, string copyName)
        {
            var copy = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads", copyName);
            File.Copy(filePath, copy, true);
            return copy;
        }

    }
}
