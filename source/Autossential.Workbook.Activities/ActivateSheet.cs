using Autossential.Shared;
using Autossential.Workbook.Activities.Extensions;
using Autossential.Workbook.Activities.Properties;
using System.Activities;

namespace Autossential.Workbook.Activities
{
    public sealed class ActivateSheet : WorkbookCodeActivity
    {
        public InArgument Sheet { get; set; }

        protected override void CacheMetadata(CodeActivityMetadata metadata)
        {
            base.CacheMetadata(metadata);

            if (Sheet == null)
            {
                metadata.AddValidationError(Resources.Validation_ValueErrorFormat(nameof(Sheet)));
            }
            else if (Sheet.IsArgumentTypeAnyCompatible<int, string>())
            {
                metadata.AddRuntimeArgument(Sheet, Sheet.ArgumentType, nameof(Sheet), true);
            }
            else
            {
                metadata.AddValidationError(Resources.Validation_TypeErrorFormat("int or string", nameof(Sheet)));
            }
        }

        protected override void Execute(CodeActivityContext context)
        {
            var sheet = Sheet.Get(context);
            if (sheet is string sheetName)
            {
                context.GetWorkbook().ActivateSheet(sheetName);
            }
            else if (sheet is int sheetIndex)
            {
                context.GetWorkbook().ActivateSheet(sheetIndex);
            }
        }
    }
}
