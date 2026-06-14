using System.Activities;

namespace Autossential.Workbook.Activities.Tests.Activities
{
    public class ReadCellTests : BaseTests
    {
        [Test]
        [Arguments(".xls", "B3")]
        [Arguments(".xlsx", "J14")]
        [Arguments(".xls", "A1")]
        public async Task ReadCell_ReturnsExpectedValue_AfterWrite(string extension, string cell)
        {
            var (processor, filePath) = NewFile(extension);
            processor.WriteCell("Sheet1", cell, "Hello");
            processor.Save();

            var value = new Variable<object>();
            var result = InvokeWorkbookScopeWith(filePath, [value], [new ReadCell {
                SheetName = "Sheet1",
                CellAddress = cell,
                Result = new OutArgument<object>(value)
            }], env => new Dictionary<string, object>
            {
                {"Value",value.Get(env) }
            });

            await Assert.That(result["Value"].ToString()).IsEqualTo("Hello");
        }
    }
}
