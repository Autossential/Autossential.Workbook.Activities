using Autossential.Shared.Tests;
using Autossential.Workbook.Activities;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Threading;

namespace Autossential.Workbook.Tests
{
    [TestClass]
    public class WriteRangeTest
    {
        [TestMethod]
        [DataRow(10, 1000)]
        [DataRow(5, 10000)]
        [DataRow(1, 100000)]
        public void OpenXML_BigTable(int cols, int rows)
        {
            var dt = CreateTable(cols, rows);

            // it will always limit the rows to 65536 

            var writeRange = new WriteRange()
            {
                AddHeaders = true,
                UseScope = false
            };

            var temp = GetTempFilePath($"OpenXML_{cols}.{rows}.xlsx");
            WorkflowTester.Invoke(writeRange, CreateArgs(temp, dt));
        }

        [TestMethod]
        public void OpenXML_TypedDataTable()
        {
            var dt = CreateTypedDataTable();
            var writeRange = new WriteRange()
            {
                AddHeaders = true,
                UseScope = false
            };

            var temp = GetTempFilePath($"OpenXML_Typed.xlsx");
            WorkflowTester.Invoke(writeRange, CreateArgs(temp, dt));
        }

        [TestMethod]
        [DataRow(10, 1000)]
        [DataRow(5, 10000)]
        [DataRow(1, 100000)]
        public void OLE2_BigTable(int cols, int rows)
        {
            var dt = CreateTable(cols, rows);

            // it will always limit the rows to 65536 

            var writeRange = new WriteRange()
            {
                AddHeaders = true,
                UseScope = false
            };

            var temp = GetTempFilePath($"OLE2_{cols}.{rows}.xls");
            WorkflowTester.Invoke(writeRange, CreateArgs(temp, dt));
        }

        [TestMethod]
        public void OLE2_TypedDataTable()
        {
            var dt = CreateTypedDataTable();
            var writeRange = new WriteRange()
            {
                AddHeaders = true,
                UseScope = false
            };

            var temp = GetTempFilePath($"OLE2_Typed.xls");
            WorkflowTester.Invoke(writeRange, CreateArgs(temp, dt));
        }

        private static DataTable CreateTypedDataTable()
        {
            var dt = new DataTable();
            dt.Columns.Add("Number1", typeof(int));
            dt.Columns.Add("Number2", typeof(double));
            dt.Columns.Add("Number3", typeof(decimal));

            dt.Columns.Add("Date", typeof(DateTime));
            dt.Columns.Add("String", typeof(string));
            dt.Columns.Add("Boolean", typeof(bool));

            var rnd = new Random();
            for (int i = 0; i < 10; i++)
            {
                var dr = dt.NewRow();
                dr["Number1"] = rnd.Next(-100, 1000);
                dr["Number2"] = (rnd.Next(-3000, 30000) * 1.5d);
                dr["Number3"] = (rnd.Next(-5000, 50000) * 2.7m);
                dr["Date"] = DateTime.Now.AddMonths(rnd.Next(0, 36)).AddDays(rnd.Next(1, 30));
                dr["String"] = Guid.NewGuid().ToString().Substring(0, 8);
                dr["Boolean"] = rnd.Next() % 2 == 0;
                dt.Rows.Add(dr);
            }
            return dt;
        }

        private static DataTable CreateTable(int cols, int rows)
        {
            var dt = new DataTable();
            for (int i = 0; i < cols; i++)
            {
                dt.Columns.Add();
            }

            for (int i = 0; i < rows; i++)
            {
                var row = dt.NewRow();
                for (int j = 0; j < cols; j++)
                {
                    row[j] = i + "." + j;
                }
                dt.Rows.Add(row);
            }

            return dt;
        }

        private static Dictionary<string, object> CreateArgs(string filePath, DataTable value)
        {
            return new Dictionary<string, object>
            {
                { nameof(WriteRange.WorkbookPath), filePath },
                { nameof(WriteRange.SheetName), "Sheet1" },
                { nameof(WriteRange.InputDataTable), value},
                { nameof(WriteRange.StartingCell), "A1" }
            };
        }

        private static string GetTempFilePath(string fileName)
        {
            return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads", fileName);
        }
    }
}
