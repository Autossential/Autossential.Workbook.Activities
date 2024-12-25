using Autossential.Workbook.Core.Internals;
using Sylvan.Data;
using Sylvan.Data.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;

namespace Autossential.Workbook.Core.Processors
{
    public abstract class WorkbookProcessorBase : IWorkbookProcessor
    {
        protected WorkbookProcessorBase(string filePath)
        {
            FilePath = filePath;

            WorkbookStream = new MemoryStream();

            var bytes = File.ReadAllBytes(FilePath);
            WorkbookStream.Write(bytes, 0, bytes.Length);
            WorkbookStream.Reset();
        }

        protected MemoryStream WorkbookStream { get; }

        public string FilePath { get; }

        public abstract void Save();

        protected bool RequiresSave { get; set; }

        private bool _disposed;

        public virtual void Dispose(bool disposing)
        {
            if (_disposed)
                return;

            if (disposing)
                WorkbookStream?.Dispose();

            _disposed = true;
        }

        public void Dispose()
        {
            try
            {
                Dispose(true);
                GC.SuppressFinalize(this);
            }
            catch (Exception e)
            {
                Trace.WriteLine(e.Message);
            }
        }

        ~WorkbookProcessorBase()
        {
            Dispose(false);
        }

        protected virtual void ValidateSheetName(string sheetName)
        {
            if (string.IsNullOrEmpty(sheetName))
                throw new ArgumentException("Sheet name cannot be null or empty", nameof(sheetName));
        }

        internal virtual RangeReference ResolveRange(string range)
        {
            if (string.IsNullOrEmpty(range))
                throw new ArgumentException("Range cannot be null or empty", nameof(range));

            var rangeRef = new OpenXMLRangeReference(range);
            if (!rangeRef.IsValid)
                throw new ArgumentException("Invalid range format", nameof(range));

            return rangeRef;
        }

        public virtual int GetColumnCount(string sheetName, string range)
        {
            ValidateSheetName(sheetName);

            var rangeRef = ResolveRange(range);
            var processedColumns = new HashSet<int>();
            var reader = GetReader();

            do
            {
                if (reader.WorksheetName != sheetName)
                    continue;

                CountColumns(reader, rangeRef, processedColumns);
                break;

            } while (reader.NextResult());

            return processedColumns.Count;
        }

        public string[] GetSheetNames()
        {
            using var reader = GetReader();
            return reader.WorksheetNames.ToArray();
        }

        public int GetRowCount(string sheetName, string range)
        {
            ValidateSheetName(sheetName);
            var rangeRef = ResolveRange(range);

            using var reader = GetReader();
            do
            {
                if (reader.WorksheetName != sheetName)
                    continue;

                if (range == "A1" || !reader.HasRows)
                    return reader.RowCount;

                return CountRows(reader, rangeRef);

            } while (reader.NextResult());

            return 0;
        }

        private ExcelDataReaderOptions GetReaderOptions(bool hasHeaders, bool useColumnDataType)
        {
            ExcelDataReaderOptions options = null;

            if (useColumnDataType)
            {
                using var r = GetReader();
                var analyzer = new SchemaAnalyzer();
                var result = analyzer.Analyze(r);
                options = new ExcelDataReaderOptions { Schema = new ExcelSchema(hasHeaders, result.GetSchema().GetColumnSchema()) };
                r.Close();
            }
            else
            {
                options = new ExcelDataReaderOptions { Schema = hasHeaders ? ExcelSchema.Dynamic : ExcelSchema.NoHeaders };
            }

            return options;
        }

        public virtual DataTable ReadRange(string sheetName, string range, bool hasHeaders, bool useColumnDataType)
        {
            ValidateSheetName(sheetName);

            var index = Array.IndexOf(GetSheetNames(), sheetName);
            if (index == -1)
                throw new ArgumentException("Sheet name not found", nameof(sheetName));

            var rangeRef = ResolveRange(range);

            using var reader = GetReader(GetReaderOptions(hasHeaders, useColumnDataType));

            var dt = new DataTable();

            do
            {
                if (reader.WorksheetName != sheetName)
                    continue;

                if (range == "A1")
                {
                    dt.Load(reader);
                    return dt;
                }

                do
                {
                    if (!rangeRef.IsRowInRange(reader.RowNumber))
                    {
                        if (reader.RowNumber > rangeRef.End.Row)
                            break;

                        continue;
                    }

                    if (dt.Columns.Count == 0)
                    {
                        rangeRef.ForEachColumn((col, index) =>
                        {
                            var value = reader.RowNumber == 1 ? reader.GetName(col - 1) : reader.GetValue(col - 1);
                            if (hasHeaders)
                            {
                                dt.Columns.Add(value?.ToString());
                            }
                            else
                            {
                                dt.Columns.Add($"Col{index + 1}");
                            }
                        });

                        if (hasHeaders)
                            continue;
                    }

                    var row = dt.NewRow();
                    rangeRef.ForEachColumn((col, index) => row[index] = reader.GetValue(col - 1));
                    dt.Rows.Add(row);

                } while (reader.Read());

            } while (reader.NextResult());

            return dt;
        }

        protected abstract ExcelDataReader GetReader(ExcelDataReaderOptions options = null);

        private static void CountColumns(ExcelDataReader reader, RangeReference rangeRef, HashSet<int> processedColumns)
        {
            var emptyColumns = new HashSet<int>();

            int colsCount = Math.Max(0, reader.FieldCount);
            if (colsCount == 0)
                colsCount = (int)rangeRef.End.Col;

            do
            {
                if (reader.RowNumber < rangeRef.Start.Row) continue;
                if (reader.RowNumber > rangeRef.End.Row) break;

                var size = Math.Min(rangeRef.End.Col, colsCount);
                for (int col = (int)rangeRef.Start.Col; col <= size; col++)
                {
                    if (processedColumns.Contains(col))
                        continue;

                    var value = reader.RowNumber == 1 ? reader.GetName(col - 1) : reader.GetValue(col - 1);

                    if (value == null || value == DBNull.Value || (value is string valueStr && string.IsNullOrEmpty(valueStr)))
                    {
                        emptyColumns.Add(col);
                        continue;
                    }

                    processedColumns.Add(col);

                    foreach (var c in emptyColumns.Where(emptyCol => emptyCol < col))
                        processedColumns.Add(c);

                    emptyColumns.RemoveWhere(c => c <= col);

                    if (processedColumns.Count == (size - (rangeRef.Start.Col - 1)))
                        return;
                }
            } while (reader.Read());
        }

        internal static int CountRows(ExcelDataReader reader, RangeReference rangeRef)
        {
            var emptyRowCount = 0;
            int count = 0;

            int colsCount = Math.Max(0, reader.FieldCount);
            if (colsCount == 0)
                colsCount = (int)rangeRef.End.Col;

            do
            {
                if (reader.RowNumber < rangeRef.Start.Row) continue;
                if (reader.RowNumber > rangeRef.End.Row) break;

                var size = Math.Min(rangeRef.End.Col, colsCount);
                for (int col = (int)rangeRef.Start.Col; col <= size; col++)
                {
                    var value = reader.RowNumber == 1 ? reader.GetName(col - 1) : reader.GetValue(col - 1);

                    if (value == null || value == DBNull.Value || (value is string valueStr && string.IsNullOrEmpty(valueStr)))
                        continue;

                    count += emptyRowCount + 1;
                    emptyRowCount = 0;
                    size = -1;

                    break;
                }

                if (size != -1)
                    emptyRowCount++;

            } while (reader.Read());

            return count;
        }

        public abstract void RenameSheet(int sheetIndex, string newSheetName);
        public abstract void RenameSheet(string fromSheetName, string toSheetName);
        public abstract void DeleteSheet(string sheetName);
        public abstract void ActivateSheet(string sheetName);
        public abstract void ActivateSheet(int sheetIndex);
        public abstract (int index, string name) GetActiveSheet();
    }
}