using Autossential.Workbook.Activities.Core;
using Xunit;

namespace Autossential.Workbook.Activities.Tests.Unit
{
    public class WorkbookProcessorTests
    {
        [Fact]
        public void TestRange()
        {
            var processor = WorkbookProcessorFactory.OpenOrCreate(@"C:\Users\alexa\Downloads\Sandbox.xlsx");
            var dt = processor.ReadRange("Sheet1", "A1", true, 3, 3);
            Assert.Equal(0, dt.Rows.Count);
        }

        [Fact]
        public void CountRows()
        {
            var processor = WorkbookProcessorFactory.OpenOrCreate(@"C:\Users\alexa\Downloads\Sandbox.xlsx");
            var count = processor.GetRowCount("Sheet1", "A1:B9");
            Assert.Equal(5, count);
        }

        [Fact]
        public void CountColumns()
        {
            var processor = WorkbookProcessorFactory.OpenOrCreate(@"C:\Users\alexa\Downloads\Sandbox.xlsx");
            var count = processor.GetColumnCount("Sheet1", "A1");
            Assert.Equal(5, count);
        }

        [Fact]
        public void ReadCell()
        {
            var processor = WorkbookProcessorFactory.OpenOrCreate(@"C:\Users\alexa\Downloads\Sandbox.xlsx", "Planilha1");
            var value = processor.ReadCell("Sheet1", "A2");
            value = processor.ReadCell("Sheet1", "C4");
        }
    }
}
