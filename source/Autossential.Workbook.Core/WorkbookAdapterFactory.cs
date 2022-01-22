using System;
using System.IO;

namespace Autossential.Workbook.Core
{
    public static class WorkbookAdapterFactory
    {
        public static IWorkbookAdapter Create(string path)
        {
            IWorkbookAdapter adapter = null;
            using (var f = File.OpenRead(path))
            {
                if (f.Length == 0)
                    throw new InvalidOperationException("The file is empty (zero bytes long)");

                var extension = Path.GetExtension(path);
                switch (extension.ToLowerInvariant())
                {
                    case ".xlsm": // macro enabled workbook
                    case ".xltm": // macro enabled template
                    case ".xlsx": // workbook
                    case ".xltx": // template
                        adapter = new OpenXmlWorkbookAdapter();
                        break;
                    default:
                        adapter = new OLE2WorkbookAdapter();
                        break;
                }
            }

            if (adapter == null)
                throw new InvalidOperationException("The file stream need be a OLE2 or OOXML stream");

            adapter.Open(path);
            return adapter;
        }
    }
}
