using Autossential.Workbook.Activities.Properties;
using Autossential.Workbook.Core.Adapters;
using System;
using System.Activities;
using System.Threading;
using System.Threading.Tasks;

namespace Autossential.Workbook.Activities
{
    public sealed class FreezeUnfreezePanes : WorkbookActivity
    {
        public InArgument<int> Cols { get; set; }
        public InArgument<int> Rows { get; set; }
        public InArgument<string> SheetName { get; set; } = "Sheet1";

        protected override void CacheMetadata(CodeActivityMetadata metadata)
        {
            base.CacheMetadata(metadata);

            if (SheetName == null)
                metadata.AddValidationError(Resources.Validation_ValueErrorFormat(nameof(SheetName)));
        }

        public override async Task<Action<AsyncCodeActivityContext>> ExecuteAsync(AsyncCodeActivityContext context, IWorkbookAdapter adapter, CancellationToken token)
        {
            var cols = Cols.Get(context);
            if (cols < 0)
                throw new ArgumentOutOfRangeException(nameof(Cols), Resources.FreezeUnfreezePanes_ErrorMsg_MinValueFormat(nameof(Cols)));

            var rows = Rows.Get(context);
            if (rows < 0)
                throw new ArgumentOutOfRangeException(nameof(Rows), Resources.FreezeUnfreezePanes_ErrorMsg_MinValueFormat(nameof(Rows)));

            var sheetName = SheetName.Get(context);
            if (string.IsNullOrEmpty(sheetName))
                throw new ArgumentException(nameof(SheetName), Resources.Validation_NullOrEmptyFormat(nameof(SheetName)));

            await Task.Run(() => adapter.FreezePanes(sheetName, cols, rows));
            return null;
        }
    }
}
