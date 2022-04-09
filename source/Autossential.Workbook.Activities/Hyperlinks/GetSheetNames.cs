using Autossential.Workbook.Core.Adapters;
using System;
using System.Activities;
using System.Threading;
using System.Threading.Tasks;

namespace Autossential.Workbook.Activities
{
    public sealed class GetSheetNames : WorkbookActivity<string[]>
    {
        public override async Task<Action<AsyncCodeActivityContext>> ExecuteAsync(AsyncCodeActivityContext context, IWorkbookAdapter adapter, CancellationToken token)
        {
            string[] names = null;
            await Task.Run(() => names = adapter.GetSheetNames());
            return ctx => Result.Set(ctx, names ?? Array.Empty<string>());
        }
    }
}
