using Autossential.Workbook.Activities.Properties;
using Autossential.Workbook.Core.Adapters;
using System;
using System.Activities;
using System.Data;
using System.Threading;
using System.Threading.Tasks;

namespace Autossential.Workbook.Activities
{
    public sealed class ReadRange : WorkbookActivity<DataTable>
    {
        public InArgument<string> SheetName { get; set; } = "Sheet1";
        public InArgument<string> Range { get; set; }
        public bool AddHeaders { get; set; } = true;

        protected override void CacheMetadata(CodeActivityMetadata metadata)
        {
            base.CacheMetadata(metadata);

            if (SheetName == null)
                metadata.AddValidationError(Resources.Validation_ValueErrorFormat(nameof(SheetName)));
        }

        public override async Task<Action<AsyncCodeActivityContext>> ExecuteAsync(AsyncCodeActivityContext context, IWorkbookAdapter adapter, CancellationToken token)
        {
            var sheetName = SheetName.Get(context);
            if (string.IsNullOrEmpty(sheetName))
                throw new ArgumentException(nameof(SheetName), Resources.Validation_NullOrEmptyFormat(nameof(SheetName)));

            var range = Range.Get(context);

            DataTable dt = null;
            await Task.Run(() => dt = adapter.ReadRange(sheetName, range, AddHeaders));
            return ctx => Result.Set(ctx, dt ?? new DataTable());
        }

    }
}