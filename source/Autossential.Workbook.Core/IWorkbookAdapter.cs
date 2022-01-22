using System;
using System.Collections.Generic;

namespace Autossential.Workbook.Core
{
    public interface IWorkbookAdapter : IDisposable
    {
        string GetSheetName(int index);
        string[] GetSheetNames();
        void AddHyperlink(string sheetName, string cellAddress, string label, string address, string tooltip);
        int RemoveHyperlink(string sheetName, string cellRange);
        IEnumerable<string> GetHyperlinks(string sheetName, string cellRange);
        void Save();
        void Open(string path);
    }
}
