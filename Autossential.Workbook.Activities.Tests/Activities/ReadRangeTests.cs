namespace Autossential.Workbook.Activities.Tests.Activities
{
    public class ReadRangeTests : BaseTests
    {
        [Test]
        [Arguments(".xlsx", "A1", true, 1, 1, 10, 10)]
        [Arguments(".xlsx", "A1", true, 3, 2, 10, 4)]
        [Arguments(".xls", "D2:G8", false, 1, 1, 4, 7)]

        public async Task ReadRange_ReturnsExpectedTable_BasedOnArguments(string extension, string range, bool hasHeaders, int headerRows, int rowsPerRecord, int expectedCols, int expectedRows)
        {
            var data = TableUtils.Build(10, 10, (col, row) =>
            {
                var value = $"C{col}R{row}";
                return col switch
                {
                    2 or 7 => null,
                    3 => row % 2 == 0 || row % 7 == 0 ? value : "",
                    8 => row < 6 ? value : null,
                    4 or 5 or 6 => (row + col) % 3 == 0 ? value : DBNull.Value,
                    10 => row > 4 && row < 10 ? value : "",
                    _ => value
                };
            });

            var (processor, filePath) = NewFile(extension);
            processor.WriteRange("Sheet1", data, "A1", hasHeaders);
            processor.Save();

            var readData = InvokeWorkbookScopeWith(filePath, new ReadRange
            {
                SheetName = "Sheet1",
                HasHeaders = hasHeaders,
                Range = range,
                RowsPerRecord = rowsPerRecord,
                HeaderRows = headerRows
            });

            await Assert.That(readData.Rows.Count).IsEqualTo(expectedRows);
            await Assert.That(readData.Columns.Count).IsEqualTo(expectedCols);
        }
    }
}
