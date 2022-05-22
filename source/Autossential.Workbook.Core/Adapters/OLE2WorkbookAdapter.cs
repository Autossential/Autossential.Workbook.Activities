using Autossential.Workbook.Core.Enums;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;

namespace Autossential.Workbook.Core.Adapters
{
    public class OLE2WorkbookAdapter : WorkbookAdapterBase
    {
        public OLE2WorkbookAdapter(string filePath) : base(filePath)
        {
        }

        public override int MaxRows => 65_536;

        public override bool IsOpenXml => true;

        private HSSFWorkbook _workbook;

        protected override IWorkbook GetWorkbook()
        {
            if (_workbook == null)
            {
                try
                {
                    WorkbookFileStream.Seek(0, SeekOrigin.Begin);
                    _workbook = new HSSFWorkbook(WorkbookFileStream);
                }
                catch (Exception e)
                {
                    Trace.WriteLine(e.Message);
                }
            }

            return _workbook;
        }

        public override void CreateNew()
        {
            using (FileStream fs = new FileStream(FilePath, FileMode.Create, FileAccess.Write))
            {
                var workbook = new HSSFWorkbook();
                workbook.CreateSheet("Sheet1");
                workbook.Write(fs);
            }
        }

        public void WriteRangeX(string sheetName, string cell, DataTable dataTable, bool addHeaders)
        {
            if (dataTable.Rows.Count == 0)
                return;

            var wb = GetWorkbook();
            var sheet = GetOrCreateSheet(sheetName);
            var cellRef = new CellReference(cell);
            var rowIndex = cellRef.Row;
            var colIndex = cellRef.Col;

            if (addHeaders)
            {
                var row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
                foreach (DataColumn col in dataTable.Columns)
                {
                    var sheetCell = row.GetCell(colIndex) ?? row.CreateCell(colIndex);
                    sheetCell.SetCellValue(col.ColumnName);

                    if (col.DataType == typeof(int))
                    {
                        var style = wb.CreateCellStyle();
                        style.DataFormat = HSSFDataFormat.GetBuiltinFormat("0");
                        sheet.SetDefaultColumnStyle(col.Ordinal, style);
                    }
                    else if (col.DataType == typeof(double) || col.DataType == typeof(decimal))
                    {
                        var style = wb.CreateCellStyle();
                        style.DataFormat = HSSFDataFormat.GetBuiltinFormat("0.00");
                        sheet.SetDefaultColumnStyle(col.Ordinal, style);
                    }
                    else if (col.DataType == typeof(DateTime))
                    {
                        var dateFormat = CultureInfo.CurrentCulture.DateTimeFormat;
                        var shortPattern = (dateFormat.ShortDatePattern + " " + dateFormat.ShortTimePattern).Replace("tt", "").Trim();
                        var style = wb.CreateCellStyle();
                        style.DataFormat = wb.CreateDataFormat().GetFormat(shortPattern);
                        sheet.SetDefaultColumnStyle(col.Ordinal, style);
                    }
                    else if (col.DataType == typeof(bool))
                    {
                        var style = wb.CreateCellStyle();
                        style.Alignment = HorizontalAlignment.Center;
                        sheet.SetDefaultColumnStyle(col.Ordinal, style);
                    }
                    else
                    {
                        var style = wb.CreateCellStyle();
                        sheet.SetDefaultColumnStyle(col.Ordinal, style);
                    }

                    colIndex++;
                }

                colIndex = cellRef.Col;
                rowIndex++;
            }

            foreach (DataRow dr in dataTable.Rows)
            {
                var row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
                for (int i = 0; i < dr.ItemArray.Length; i++)
                {
                    var sheetCell = row.GetCell(colIndex + i) ?? row.CreateCell(colIndex + i);
                    SetCellValue(sheetCell, dr[i]);
                }

                if (++rowIndex == MaxRows)
                    break;
            }

            foreach (DataColumn col in dataTable.Columns)
                sheet.AutoSizeColumn(col.Ordinal);

            RequiresSave();
        }

        public override void Dispose(bool disposing)
        {
            if (disposing && _workbook != null)
                _workbook.Close();
        }

        private HSSFPalette _palette;
        protected override IColor ConvertColor(Color color)
        {
            _palette = _palette ?? ((HSSFWorkbook)GetWorkbook()).GetCustomPalette();
            return _palette.FindSimilarColor(color.R, color.G, color.B);
        }

        public override void DrawBorder(string sheetName, string range, Border border, Enums.BorderStyle style, Color color)
        {
            var wb = GetWorkbook();
            var cellStyle = wb.CreateCellStyle();

            _ = GetOrCreateSheet(sheetName);
            var cells = GetCells(sheetName, range);

            var borderStyle = ConvertBorderStyle(style);
            var borderColor = ConvertColor(color).Indexed;

            switch (border)
            {
                case Border.None:
                case Border.All:

                    cellStyle.BorderLeft = borderStyle;
                    cellStyle.BorderTop = borderStyle;
                    cellStyle.BorderRight = borderStyle;
                    cellStyle.BorderBottom = borderStyle;

                    cellStyle.LeftBorderColor = borderColor;
                    cellStyle.TopBorderColor = borderColor;
                    cellStyle.RightBorderColor = borderColor;
                    cellStyle.BottomBorderColor = borderColor;

                    foreach (var cell in cells)
                        cell.CellStyle = cellStyle;

                    break;
                case Border.Outside:
                    break;
                case Border.Inside:
                    break;
                case Border.Top:
                    cellStyle.BorderTop = borderStyle;
                    break;
                case Border.Bottom:
                    cellStyle.BorderBottom = borderStyle;
                    break;
                case Border.Left:
                    cellStyle.BorderLeft = borderStyle;
                    break;
                case Border.Right:
                    cellStyle.BorderRight = borderStyle;
                    break;
            }

            RequiresSave();
        }
    }


    //    private NPOI.HSSF.Util.HSSFColor GetDarkerColor(NPOI.HSSF.Util.HSSFColor xlsColor, HSSFPalette palette)
    //    {
    //        var rgb = xlsColor.RGB;
    //        while (rgb[0] + rgb[1] + rgb[2] > 0)
    //        {
    //            if (rgb[0] > 0) rgb[0]--;
    //            if (rgb[1] > 0) rgb[1]--;
    //            if (rgb[2] > 0) rgb[2]--;

    //            var dark = palette.FindSimilarColor(rgb[0], rgb[1], rgb[2]);
    //            if (!dark.RGB.SequenceEqual(xlsColor.RGB))
    //                return dark;
    //        }
    //        return xlsColor;
    //    }

    //    public override void DrawBorder(string sheetName, string range, Border border, Enums.BorderStyle style, Color color)
    //    {
    //        throw new NotImplementedException();
    //    }

    //    public override void FillColor(string sheetName, string range, Color[] colors, FillOrientation orientation)
    //    {
    //        var wb = (HSSFWorkbook)GetWorkbook();
    //        var sheet = GetOrCreateSheet(sheetName);
    //        var palette = wb.GetCustomPalette();
    //        var len = colors.Length;
    //        var cells = GetCells(sheetName, range);

    //        if (!cells.Any() || len == 0)
    //            return;

    //        var index = 0;


    //        if (orientation == FillOrientation.Chess)
    //        {
    //            foreach (var cell in cells)
    //            {
    //                cell.CellStyle = styles[index];
    //                if (++index == len)
    //                    index = 0;
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
    //}
}