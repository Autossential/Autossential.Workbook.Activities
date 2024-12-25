using Autossential.Workbook.Activities.Extensions;
using System.Activities;

namespace Autossential.Workbook.Activities
{
    public sealed class GetActiveSheet : WorkbookCodeActivity
    {
        public OutArgument<int> SheetIndex { get; set; }
        public OutArgument<string> SheetName { get; set; }
        protected override void Execute(CodeActivityContext context)
        {
            var (index, name) = context.GetWorkbook().GetActiveSheet();

            SheetIndex.Set(context, index);
            SheetName.Set(context, name);
        }
    }
}
