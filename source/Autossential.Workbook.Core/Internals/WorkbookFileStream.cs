using System.IO;

namespace Autossential.Workbook.Core.Internals
{
    public sealed class WorkbookFileStream : FileStream
    {
        public WorkbookFileStream(string path, FileMode mode) : base(path, mode)
        {
        }

        public override void Close()
        {
            // do not close.
        }

        public void CloseWorkbook()
        {
            base.Close();
        }
    }

}
