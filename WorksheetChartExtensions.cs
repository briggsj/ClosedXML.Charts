using System.IO;
using System;

using ClosedXML.Excel;

namespace ClosedXML.Charts
{
    public static class WorksheetChartExtensions
    {
        /// <summary>
        /// Convenience method that creates a chart after building the ClosedXML workbook.
        /// - workbook: ClosedXML XLWorkbook populated.
        /// - sheetName: worksheet name to place the chart on.
        /// - categoryRange: e.g. "A2:A6"
        /// - valuesRange: e.g. "B2:B6"
        /// This saves the workbook into a MemoryStream, injects the chart, and returns the stream repositioned to 0.
        /// Caller is responsible for writing stream to file or further processing.
        /// </summary>
        public static MemoryStream SaveWithBarChart(this XLWorkbook workbook, string sheetName,
            string categoryRange, string valuesRange, string chartTitle = "Chart")
        {
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));
            var ms = new MemoryStream();
            workbook.SaveAs(ms);
            // ChartHelper will operate on MS directly
            ChartHelper.AddBarChartToWorkbookStream(ms, sheetName, categoryRange, valuesRange, chartTitle);
            // Reset position for caller
            ms.Position = 0;
            return ms;
        }
    }
}