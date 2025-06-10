using System;
using System.Collections.ObjectModel;
using System.Data;
using System.Data.Common;

namespace Autossential.Workbook.Core.Extensions
{
    internal static class DataTableExtensions
    {
        public static void AddColumns(this DataTable dt, int startCol, int endCol, Func<int, int, string> columnNameHandler, ReadOnlyCollection<DbColumn> columnSchema)
        {
            for (int i = 0, j = startCol, k = j - 1; j <= endCol; i++, j++, k = j - 1)
            {
                if (k < columnSchema.Count)
                {
                    dt.Columns.Add(columnNameHandler(k, i), columnSchema[k].DataType);
                    continue;
                }

                dt.Columns.Add(columnNameHandler(k, i), typeof(object));
            }
        }
    }
}
