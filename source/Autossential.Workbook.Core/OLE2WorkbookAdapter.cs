using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;

namespace Autossential.Workbook.Core
{
    public class OLE2WorkbookAdapter : IWorkbookAdapter
    {
        private HSSFWorkbook _workbook;
        private string _filePath;

        public void AddHyperlink(string sheetName, string cellAddress, string label, string address, string tooltip)
        {
            var sheet = _workbook.GetSheet(sheetName);
            var cellRef = new CellReference(cellAddress);
            var row = sheet.GetRow(cellRef.Row) ?? sheet.CreateRow(cellRef.Row);
            var cell = row.GetCell(cellRef.Col) ?? row.CreateCell(cellRef.Col);

            var linkType = HyperlinkType.Document;

            if (Regex.IsMatch(address, "(https?|ftp)://", RegexOptions.IgnoreCase))
            {
                linkType = HyperlinkType.Url;
            }
            else if (address.StartsWith("mailto:", StringComparison.OrdinalIgnoreCase))
            {
                linkType = HyperlinkType.Email;
            }
            else if (!string.IsNullOrEmpty(Path.GetExtension(address)))
            {
                linkType = HyperlinkType.File;
            }

            cell.Hyperlink = new HSSFHyperlink(linkType) { Address = address };
            if (string.IsNullOrEmpty(label))
                label = address;

            cell.SetCellValue(label);
            Save();
        }

        public void Dispose()
        {
            _workbook.Close();
        }

        public IEnumerable<string> GetHyperlinks(string sheetName, string cellRange)
        {
            foreach (var cell in GetCells(sheetName, cellRange))
            {
                if (cell.Hyperlink != null)
                    yield return cell.Hyperlink.Address;
            }
        }

        public string GetSheetName(int index)
        {
            return _workbook.GetSheetName(index);
        }

        public string[] GetSheetNames()
        {
            var names = new string[_workbook.NumberOfSheets];
            for (int i = 0; i < names.Length; i++)
                names[i] = GetSheetName(i);

            return names;
        }

        public void Open(string path)
        {
            _workbook = new HSSFWorkbook(File.OpenRead(path));
            _filePath = path;
        }

        public int RemoveHyperlink(string sheetName, string cellRange)
        {
            int count = 0;
            foreach (var cell in GetCells(sheetName, cellRange))
            {
                if (cell.Hyperlink == null)
                    continue;

                count++;
                cell.RemoveHyperlink();
            }

            Save();
            return count;
        }

        public void Save()
        {
            using (FileStream fs = new FileStream(_filePath, FileMode.Create, FileAccess.Write))
                _workbook.Write(fs);
        }

        private IEnumerable<ICell> GetCells(string sheetName, string cellRange)
        {
            var sheet = _workbook.GetSheet(sheetName);
            if (string.IsNullOrEmpty(cellRange))
            {
                for (int i = sheet.FirstRowNum; i <= sheet.LastRowNum; i++)
                {
                    var row = sheet.GetRow(i);
                    if (row == null)
                        continue;

                    for (int j = row.FirstCellNum; j < row.LastCellNum; j++)
                    {
                        var cell = row.GetCell(j);
                        if (cell != null)
                            yield return cell;
                    }
                }
            }
            else
            {
                var range = CellRangeAddress.ValueOf(cellRange);

                var firstRow = range.FirstRow == -1 ? sheet.FirstRowNum : range.FirstRow;
                var lastRow = range.LastRow == -1 ? sheet.LastRowNum : range.LastRow;

                for (int i = firstRow; i <= lastRow; i++)
                {
                    var row = sheet.GetRow(i);
                    if (row == null)
                        continue;

                    for (int j = range.FirstColumn; j <= range.LastColumn; j++)
                    {
                        var cell = row.GetCell(j);
                        if (cell != null)
                            yield return cell;
                    }
                }
            }
        }
    }
}
