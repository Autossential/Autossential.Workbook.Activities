using Autossential.Workbook.Activities.Core;
using System.Activities;

namespace Autossential.Workbook.Activities.Tests.Activities
{
    internal class WriteCellTests : BaseTests
    {
        [Test]
        [Arguments(".xlsx", "A1", "Hello")]
        [Arguments(".xls", "B2", 1)]
        public async Task WriteCell_CellIsUpdated_BasedOnArguments(string extension, string cell, object value)
        {
            object readValue = Run(extension, cell, value);

            await Assert.That(value?.ToString()).IsEqualTo(readValue?.ToString());
        }

        [Test]
        [Arguments(".xlsx", null, "Hello")]
        [Arguments(".xls", "", 1)]
        public async Task WriteCell_Fails_WhenMissingCell(string extension, string? cell, object value)
        {
            object readValue = Run(extension, cell, value);

            await Assert.That(value?.ToString()).IsEqualTo(readValue?.ToString());
        }

        private object Run(string extension, string? cell, object value)
        {
            var filePath = NewTempFilePath(extension);
            InvokeWorkbookScopeWith(filePath, new WriteCell
            {
                SheetName = "Sheet1",
                Cell = cell,
                Value = new InArgument<object>(ctx => value)
            });

            var processor = WorkbookProcessorFactory.OpenOrCreate(filePath);
            var readValue = processor.ReadCell("Sheet1", cell);
            return readValue;
        }
    }
}