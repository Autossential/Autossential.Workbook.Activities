using Autossential.Workbook.Activities.Core;
using Xunit;

namespace Autossential.Workbook.Activities.Tests.Unit
{
    public class WorkbookProcessorTests
    {
        private string filePath = @"C:\Users\alexa\Downloads\Sandbox_copy.xlsx";
        public WorkbookProcessorTests()
        {
            if (!File.Exists(filePath))
                File.Copy(@"C:\Users\alexa\Downloads\Sandbox.xlsx", filePath);
        }

        [Theory]
        [InlineData("Empty", "B3:E9", 7, 4, false)]
        [InlineData("Empty", "B3:E9", 6, 4, true)]
        [InlineData("Empty", "B3", 0, 0, false)]
        [InlineData("Empty", "", 0, 0, true)]

        [InlineData("Companies", "", 11, 7, false)]
        [InlineData("Companies", "A1", 10, 7, true)]
        [InlineData("Companies", "C1", 10, 5, true)]

        [InlineData("Animals", "A1", 8, 4, true)]
        [InlineData("Animals", "B4", 5, 3, true)]

        [InlineData("Cars", "A1", 8, 4, false)]
        [InlineData("Cars", "B3:F11", 8, 5, true)]
        [InlineData("Cars", "A1:E10", 10, 5, false)]
        public void ReadRange(string sheetName, string range, int rows, int cols, bool hasHeaders)
        {
            var processor = WorkbookProcessorFactory.OpenOrCreate(@"C:\Users\alexa\Downloads\Sandbox_copy.xlsx");
            var table = processor.ReadRange(sheetName, range, hasHeaders, 1, 1);
            Assert.Equal(rows, table.Rows.Count);
            Assert.Equal(cols, table.Columns.Count);
        }
    }
}
