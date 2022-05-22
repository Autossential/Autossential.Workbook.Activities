using Autossential.Workbook.Core.Enums;
using ClosedXML.Excel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.Streaming;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;

namespace Autossential.Workbook.Core.Adapters
{
    public class OpenXmlWorkbookAdapter : WorkbookAdapterBase
    {
        private XSSFWorkbook _workbook;

        public OpenXmlWorkbookAdapter(string filePath) : base(filePath)
        {
        }

        public override int MaxRows => 1_048_576;

        public override bool IsOpenXml => true;

        public override void CreateNew()
        {
            using (var fs = new FileStream(FilePath, FileMode.Create, FileAccess.Write))
            {
                var workbook = new XSSFWorkbook();
                workbook.CreateSheet("Sheet1");
                workbook.Write(fs);
            }
        }

        public override void Dispose(bool disposing)
        {
            if (disposing && _workbook != null)
                _workbook.Close();
        }

        private static XSSFCellStyle CreateBorderStyle(IWorkbook workbook, string anchors, NPOI.SS.UserModel.BorderStyle borderStyle, XSSFColor borderColor)
        {
            var style = (XSSFCellStyle)workbook.CreateCellStyle();
            if (anchors.Contains('T'))
            {
                style.BorderTop = borderStyle;
                style.SetTopBorderColor(borderColor);
            }
            if (anchors.Contains('B'))
            {
                style.BorderBottom = borderStyle;
                style.SetBottomBorderColor(borderColor);
            }
            if (anchors.Contains('L'))
            {
                style.BorderLeft = borderStyle;
                style.SetLeftBorderColor(borderColor);
            }
            if (anchors.Contains('R'))
            {
                style.BorderRight = borderStyle;
                style.SetRightBorderColor(borderColor);
            }
            return style;
        }

        public override void DrawBorder(string sheetName, string range, Border border, Enums.BorderStyle style, Color color)
        {
            var wb = GetWorkbook();

            _ = GetOrCreateSheet(sheetName);
            var cells = GetCells(sheetName, range);

            var borderStyle = ConvertBorderStyle(style);
            var borderColor = (XSSFColor)ConvertColor(color);

            if (border == Border.Outside)
            {
                var first = cells.First();
                var last = cells.Last();

                var tlStyle = CreateBorderStyle(wb, "TL", borderStyle, borderColor);
                var trStyle = CreateBorderStyle(wb, "TR", borderStyle, borderColor);
                var blStyle = CreateBorderStyle(wb, "BL", borderStyle, borderColor);
                var brStyle = CreateBorderStyle(wb, "BR", borderStyle, borderColor);

                first.CellStyle = tlStyle;
                last.CellStyle = brStyle;

                cells.First(p => p.RowIndex == first.RowIndex && p.ColumnIndex == last.ColumnIndex).CellStyle = trStyle;
                cells.First(p => p.RowIndex == last.RowIndex && p.ColumnIndex == first.ColumnIndex).CellStyle = blStyle;

                var topStyle = CreateBorderStyle(wb, "T", borderStyle, borderColor);
                var bottomStyle = CreateBorderStyle(wb, "B", borderStyle, borderColor);
                var leftStyle = CreateBorderStyle(wb, "L", borderStyle, borderColor);
                var rightStyle = CreateBorderStyle(wb, "R", borderStyle, borderColor);

                foreach (var cell in cells)
                {
                    if (cell.RowIndex == first.RowIndex && cell.ColumnIndex > first.ColumnIndex && cell.ColumnIndex < last.ColumnIndex)
                    {
                        cell.CellStyle = topStyle;
                    }
                    else if (cell.RowIndex == last.RowIndex && cell.ColumnIndex > first.ColumnIndex && cell.ColumnIndex < last.ColumnIndex)
                    {
                        cell.CellStyle = bottomStyle;
                    }
                    else if (cell.ColumnIndex == first.ColumnIndex && cell.RowIndex > first.RowIndex && cell.RowIndex < last.RowIndex)
                    {
                        cell.CellStyle = leftStyle;
                    }
                    else if (cell.ColumnIndex == last.ColumnIndex && cell.RowIndex > first.RowIndex && cell.RowIndex < last.RowIndex)
                    {
                        cell.CellStyle = rightStyle;
                    }
                }

                RequiresSave();

                return;
            }

            ICellStyle cellStyle = null;
            switch (border)
            {
                case Border.None:
                case Border.All:
                    cellStyle = CreateBorderStyle(wb, "TBLR", borderStyle, borderColor);
                    break;
                case Border.Top:
                    cellStyle = CreateBorderStyle(wb, "T", borderStyle, borderColor);
                    break;
                case Border.Bottom:
                    cellStyle = CreateBorderStyle(wb, "B", borderStyle, borderColor);
                    break;
                case Border.Left:
                    cellStyle = CreateBorderStyle(wb, "L", borderStyle, borderColor);
                    break;
                case Border.Right:
                    cellStyle = CreateBorderStyle(wb, "R", borderStyle, borderColor);
                    break;
            }

            foreach (var cell in cells)
                cell.CellStyle = cellStyle;

            RequiresSave();
        }

        protected override IColor ConvertColor(Color color)
        {
            return new XSSFColor(color);
        }

        protected override IWorkbook GetWorkbook()
        {
            if (_workbook == null)
            {
                try
                {
                    WorkbookFileStream.Seek(0, SeekOrigin.Begin);
                    _workbook = new XSSFWorkbook(WorkbookFileStream);
                }
                catch (Exception e)
                {
                    Trace.WriteLine(e.Message);
                }
            }

            return _workbook;
        }
    }
    //public class OpenXmlWorkbookAdapter : WorkbookAdapterBase
    //{
    //    private const int MAX_ROWS = 1_048_576;
    //    private XSSFWorkbook _workbook = null;

    //    public OpenXmlWorkbookAdapter(string filePath) : base(filePath)
    //    {
    //    }
    //    protected override IWorkbook GetWorkbook()
    //    {
    //        
    //        
    //        
    //        
    //        
    //        
    //        
    //        
    //        
    //        
    //        
    //        

    //        
    //    }

    //    public override void CreateNew()
    //    {
    //        throw new NotImplementedException();
    //    }

    //    public override void Dispose(bool disposing)
    //    {
    //       
    //       
    //    }

    //    public override void DrawBorder(string sheetName, string range, Border border, Enums.BorderStyle style, Color color)
    //    {
    //        throw new NotImplementedException();
    //    }

    //    public override void FillColor(string sheetName, string range, Color[] colors, FillOrientation orientation)
    //    {
    //        var wb = GetWorkbook();
    //        var sheet = GetOrCreateSheet(sheetName);
    //        var len = colors.Length;
    //        var cells = GetCells(sheetName, range);

    //        if (!cells.Any() || len == 0)
    //            return;

    //        var index = 0;

    //        var styles = colors.Select(c =>
    //        {
    //            var style = (XSSFCellStyle)wb.CreateCellStyle();
    //            style.SetFillForegroundColor(new XSSFColor(c));
    //            style.FillPattern = FillPattern.SolidForeground;
    //            return style;
    //        }).ToArray();

    //        if (orientation == FillOrientation.Chess)
    //        {
    //            var lastRow = cells.First().RowIndex;
    //            var colsCount = 0;
    //            foreach (var cell in cells)
    //            {
    //                if (lastRow != cell.RowIndex)
    //                {
    //                    if (len == colsCount || colsCount % len == 0)
    //                        index--;

    //                    lastRow = cell.RowIndex;
    //                    colsCount = 0;
    //                }

    //                if (index == len)
    //                    index = 0;

    //                cell.CellStyle = styles[index];

    //                index++;
    //                colsCount++;
    //            }
    //        }
    //        else
    //        {
    //            var firstCell = cells.First();
    //            int currentCol, currentRow, firstCol;

    //            switch (orientation)
    //            {
    //                case FillOrientation.Horizontal:

    //                    currentRow = firstCell.RowIndex;
    //                    foreach (var cell in cells)
    //                    {
    //                        if (currentRow != cell.RowIndex)
    //                        {
    //                            currentRow = cell.RowIndex;
    //                            if (++index == len)
    //                                index = 0;
    //                        }

    //                        cell.CellStyle = styles[index];
    //                    }

    //                    break;

    //                case FillOrientation.Vertical:

    //                    firstCol = firstCell.ColumnIndex;
    //                    currentCol = firstCol;

    //                    foreach (var cell in cells)
    //                    {
    //                        if (currentCol != cell.ColumnIndex)
    //                        {
    //                            currentCol = cell.ColumnIndex;
    //                            if (++index == len || currentCol == firstCol)
    //                                index = 0;
    //                        }

    //                        cell.CellStyle = styles[index];
    //                    }

    //                    break;
    //                case FillOrientation.DiagonalLeft:

    //                    currentCol = firstCell.ColumnIndex;
    //                    currentRow = firstCell.RowIndex;

    //                    foreach (var cell in cells)
    //                    {
    //                        if (currentRow != cell.RowIndex)
    //                        {
    //                            currentCol++;
    //                            currentRow = cell.RowIndex;
    //                        }

    //                        if (currentCol == cell.ColumnIndex)
    //                        {
    //                            cell.CellStyle = styles[index];
    //                            if (++index == len)
    //                                index = 0;
    //                        }
    //                    }

    //                    break;
    //                case FillOrientation.DiagonalRight:

    //                    currentCol = cells.Last().ColumnIndex;
    //                    currentRow = firstCell.RowIndex;

    //                    foreach (var cell in cells)
    //                    {
    //                        if (currentRow != cell.RowIndex)
    //                        {
    //                            currentCol--;
    //                            currentRow = cell.RowIndex;
    //                        }

    //                        if (currentCol == cell.ColumnIndex)
    //                        {
    //                            cell.CellStyle = styles[index];
    //                            if (++index == len)
    //                                index = 0;
    //                        }
    //                    }

    //                    break;
    //            }
    //        }

    //        RequiresSave();
    //    }

    //    public override Action GetSaveHandler() => () =>
    //    {
    //        using (FileStream fs = new FileStream(FilePath, FileMode.Create, FileAccess.Write))
    //        {
    //            GetWorkbook().Write(fs);
    //        }
    //    };

    //    public override void WriteRange(string sheetName, string cell, DataTable dataTable, bool addHeaders)
    //    {
    //        if (dataTable.Rows.Count == 0)
    //            return;

    //        var sheet = GetOrCreateSheet(sheetName);
    //        var cellRef = new CellReference(cell);
    //        var rowIndex = cellRef.Row;
    //        var colIndex = cellRef.Col;

    //        if (addHeaders)
    //        {
    //            var row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
    //            var format = _workbook.CreateDataFormat();
    //            foreach (DataColumn col in dataTable.Columns)
    //            {
    //                var sheetCell = row.GetCell(colIndex) ?? row.CreateCell(colIndex);
    //                sheetCell.SetCellValue(col.ColumnName);

    //                if (col.DataType == typeof(int))
    //                {
    //                    var style = _workbook.CreateCellStyle();
    //                    style.DataFormat = format.GetFormat("0");
    //                    sheet.SetDefaultColumnStyle(col.Ordinal, style);
    //                }
    //                else if (col.DataType == typeof(double) || col.DataType == typeof(decimal))
    //                {
    //                    var style = _workbook.CreateCellStyle();
    //                    style.DataFormat = format.GetFormat("0.00"); 
    //                    sheet.SetDefaultColumnStyle(col.Ordinal, style);
    //                }
    //                else if (col.DataType == typeof(DateTime))
    //                {
    //                    var dateFormat = CultureInfo.CurrentCulture.DateTimeFormat;
    //                    var shortPattern = (dateFormat.ShortDatePattern + " " + dateFormat.ShortTimePattern).Replace("tt", "").Trim();
    //                    var style = _workbook.CreateCellStyle();
    //                    style.DataFormat = format.GetFormat(shortPattern);
    //                    sheet.SetDefaultColumnStyle(col.Ordinal, style);
    //                }
    //                else if (col.DataType == typeof(bool))
    //                {
    //                    var style = _workbook.CreateCellStyle();
    //                    style.Alignment = HorizontalAlignment.Center;
    //                    sheet.SetDefaultColumnStyle(col.Ordinal, style);
    //                }
    //                else
    //                {
    //                    var style = _workbook.CreateCellStyle();
    //                    sheet.SetDefaultColumnStyle(col.Ordinal, style);
    //                }

    //                colIndex++;
    //            }

    //            colIndex = cellRef.Col;
    //            rowIndex++;
    //        }

    //        foreach (DataRow dr in dataTable.Rows)
    //        {
    //            var row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
    //            for (int i = 0; i < dr.ItemArray.Length; i++)
    //            {
    //                var sheetCell = row.GetCell(colIndex + i) ?? row.CreateCell(colIndex + i);
    //                SetCellValue(sheetCell, dr[i]);
    //            }

    //            if (++rowIndex == MAX_ROWS)
    //                break;
    //        }

    //        foreach (DataColumn col in dataTable.Columns)
    //            sheet.AutoSizeColumn(col.Ordinal);

    //        RequiresSave();
    //    }
    //}
}