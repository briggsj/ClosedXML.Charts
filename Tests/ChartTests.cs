using Microsoft.VisualStudio.TestTools.UnitTesting;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using System;
using System.IO;
using System.Linq;

namespace ClosedXML.Charts.Tests
{
    [TestClass]
    public class ChartHelperTests
    {
        [TestMethod]
        public void AddBarChartToWorkbookStream_ValidInputs_AddsChart()
        {
            // Arrange
            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet1");
            ws.Cell("A1").Value = "Category";
            ws.Cell("B1").Value = "Value";
            ws.Cell("A2").Value = "A";
            ws.Cell("B2").Value = 1;
            ws.Cell("A3").Value = "B";
            ws.Cell("B3").Value = 2;
            var ms = new MemoryStream();
            wb.SaveAs(ms);

            // Act
            ChartHelper.AddBarChartToWorkbookStream(ms, "Sheet1", "A2:A3", "B2:B3", "Test Chart");

            // Assert
            ms.Position = 0;
            using var doc = SpreadsheetDocument.Open(ms, false);
            var workbookPart = doc.WorkbookPart;
            var sheet = workbookPart.Workbook.Descendants<DocumentFormat.OpenXml.Spreadsheet.Sheet>().First(s => s.Name.Value == "Sheet1");
            var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
            Assert.IsNotNull(worksheetPart.DrawingsPart, "DrawingsPart should be created");
            Assert.IsTrue(worksheetPart.DrawingsPart.ChartParts.Any(), "At least one ChartPart should exist");
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentNullException))]
        public void AddBarChartToWorkbookStream_NullStream_ThrowsArgumentNullException()
        {
            // Act
            ChartHelper.AddBarChartToWorkbookStream(null, "Sheet1", "A1:A2", "B1:B2");
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void AddBarChartToWorkbookStream_NonSeekableStream_ThrowsArgumentException()
        {
            // Arrange
            var nonSeekableStream = new NonSeekableStream();

            // Act
            ChartHelper.AddBarChartToWorkbookStream(nonSeekableStream, "Sheet1", "A1:A2", "B1:B2");
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void AddBarChartToWorkbookStream_SheetNotFound_ThrowsArgumentException()
        {
            // Arrange
            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet1");
            var ms = new MemoryStream();
            wb.SaveAs(ms);

            // Act
            ChartHelper.AddBarChartToWorkbookStream(ms, "NonExistentSheet", "A1:A2", "B1:B2");
        }

        private class NonSeekableStream : MemoryStream
        {
            public override bool CanSeek => false;
        }
    }

    [TestClass]
    public class WorksheetChartExtensionsTests
    {
        [TestMethod]
        public void SaveWithBarChart_ValidInputs_ReturnsStreamWithChart()
        {
            // Arrange
            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet1");
            ws.Cell("A1").Value = "Category";
            ws.Cell("B1").Value = "Value";
            ws.Cell("A2").Value = "A";
            ws.Cell("B2").Value = 1;
            ws.Cell("A3").Value = "B";
            ws.Cell("B3").Value = 2;

            // Act
            var ms = wb.SaveWithBarChart("Sheet1", "A2:A3", "B2:B3", "Test Chart");

            // Assert
            using var doc = SpreadsheetDocument.Open(ms, false);
            var workbookPart = doc.WorkbookPart;
            var sheet = workbookPart.Workbook.Descendants<DocumentFormat.OpenXml.Spreadsheet.Sheet>().First(s => s.Name.Value == "Sheet1");
            var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
            Assert.IsNotNull(worksheetPart.DrawingsPart, "DrawingsPart should be created");
            Assert.IsTrue(worksheetPart.DrawingsPart.ChartParts.Any(), "At least one ChartPart should exist");
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentNullException))]
        public void SaveWithBarChart_NullWorkbook_ThrowsArgumentNullException()
        {
            // Act
            XLWorkbook nullWb = null;
            nullWb.SaveWithBarChart("Sheet1", "A1:A2", "B1:B2");
        }
    }
}