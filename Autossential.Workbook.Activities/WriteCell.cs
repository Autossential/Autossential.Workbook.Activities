using Autossential.Workbook.Activities.Base;
using Autossential.Workbook.Activities.Extensions;
using System.Activities;

namespace Autossential.Workbook.Activities
{
    public sealed class WriteCell : WorkbookCodeActivity
    {
        [RequiredArgument]
        public InArgument<string> SheetName { get; set; }
        
        [RequiredArgument]
        public InArgument<string> Cell { get; set; }

        [RequiredArgument]
        public InArgument<object> Value { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            var sheetName = SheetName.Get(context);
            var cell = Cell.Get(context);
            var value = Value.Get(context);

            context.GetWorkbookProcessor().WriteCell(sheetName, cell, value);
        }
    }
}
