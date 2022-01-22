using ClosedXML.Excel;
using System.Collections.Generic;

namespace Autossential.Workbook.Core
{
    public class OpenXmlWorkbookAdapter : IWorkbookAdapter
    {
        private XLWorkbook _workbook;

        public void AddHyperlink(string sheetName, string cellAddress, string label, string address, string tooltip)
        {
            var sheet = _workbook.Worksheet(sheetName) ?? _workbook.AddWorksheet(sheetName);
            var cell = sheet.Cell(cellAddress);
            cell.Hyperlink = new XLHyperlink(address, tooltip);
            if (string.IsNullOrEmpty(label))
                label = address;

            cell.SetValue(label);
            Save();
        }

        public void Dispose()
        {
            _workbook?.Dispose();
        }

        public IEnumerable<string> GetHyperlinks(string sheetName, string cellRange)
        {
            foreach (var cell in GetCells(sheetName, cellRange))
            {
                if (cell.HasHyperlink)
                {
                    if (cell.Hyperlink.IsExternal)
                    {
                        yield return cell.Hyperlink.ExternalAddress.ToString();
                        continue;
                    }

                    yield return cell.Hyperlink.InternalAddress.ToString();
                }
            }
        }

        public string GetSheetName(int index)
        {
            return _workbook.Worksheet(index + 1)?.Name;
        }

        public string[] GetSheetNames()
        {
            var names = new string[_workbook.Worksheets.Count];
            for (int i = 0; i < names.Length; i++)
                names[i] = _workbook.Worksheet(i + 1).Name;

            return names;
        }

        public void Open(string path)
        {
            _workbook = new XLWorkbook(path);
        }

        public int RemoveHyperlink(string sheetName, string cellRange)
        {
            var count = 0;
            foreach (var cell in GetCells(sheetName, cellRange))
            {
                if (cell.HasHyperlink)
                {
                    cell.Hyperlink.Delete();
                    count++;
                }
            }

            Save();
            return count;
        }

        private IXLCells GetCells(string sheetName, string cellRange)
        {
            var sheet = _workbook.Worksheet(sheetName);
            return string.IsNullOrEmpty(cellRange) ? sheet.CellsUsed() : sheet.Range(cellRange).CellsUsed();
        }

        public void Save()
        {
            _workbook.Save();
        }
    }
}
