using System;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Drawing.Spreadsheet;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace ClosedXML.Charts
{
    public static class ChartHelper
    {
        /// <summary>
        /// Adds a simple clustered column (bar) chart to a worksheet inside an existing workbook stream.
        /// The workbookStream must be seekable and opened for read/write (e.g. MemoryStream).
        /// categoryRange and valuesRange should be Excel address strings like "A2:A6" and sheetName must match.
        /// </summary>
        public static void AddBarChartToWorkbookStream(Stream workbookStream, string sheetName,
            string categoryRangeAddress, string valuesRangeAddress, string chartTitle = "Chart")
        {
            if (workbookStream == null) throw new ArgumentNullException(nameof(workbookStream));
            if (!workbookStream.CanSeek || !workbookStream.CanRead || !workbookStream.CanWrite)
                throw new ArgumentException("workbookStream must be seekable and opened for read/write (e.g. MemoryStream).");

            workbookStream.Position = 0;

            using (var spreadsheet = SpreadsheetDocument.Open(workbookStream, true))
            {
                var workbookPart = spreadsheet.WorkbookPart;
                if (workbookPart == null) throw new InvalidOperationException("WorkbookPart missing.");

                // Find sheet by name
                Sheet sheet = null;
                foreach (var s in workbookPart.Workbook.Descendants<Sheet>())
                {
                    if (string.Equals(s.Name.Value, sheetName, StringComparison.OrdinalIgnoreCase))
                    {
                        sheet = s;
                        break;
                    }
                }
                if (sheet == null) throw new ArgumentException($"Sheet '{sheetName}' not found in workbook.");

                var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);

                // Ensure a DrawingsPart exists
                DrawingsPart drawingsPart;
                if (worksheetPart.DrawingsPart == null)
                {
                    drawingsPart = worksheetPart.AddNewPart<DrawingsPart>();
                    drawingsPart.WorksheetDrawing = new WorksheetDrawing();
                    drawingsPart.WorksheetDrawing.Save();

                    // create relationship in worksheet
                    var drawings = new DocumentFormat.OpenXml.Spreadsheet.Drawing() { Id = worksheetPart.GetIdOfPart(drawingsPart) };
                    var ws = worksheetPart.Worksheet;
                    ws.Append(drawings);
                    ws.Save();
                }
                else
                {
                    drawingsPart = worksheetPart.DrawingsPart;
                }

                // Add ChartPart
                var chartPart = drawingsPart.AddNewPart<ChartPart>();
                GenerateChartPartContent(chartPart, sheetName, categoryRangeAddress, valuesRangeAddress, chartTitle);

                // Add a GraphicFrame in the drawing with a ChartReference
                var chartRelId = drawingsPart.GetIdOfPart(chartPart);

                // Create a TwoCellAnchor to position the chart; coordinates are basic placeholders
                var twoCell = new TwoCellAnchor(
                    new DocumentFormat.OpenXml.Spreadsheet.FromMarker(
                        new ColumnId("0"),
                        new ColumnOffset("0"),
                        new RowId("1"),
                        new RowOffset("0")
                    ),
                    new DocumentFormat.OpenXml.Spreadsheet.ToMarker(
                        new ColumnId("8"),
                        new ColumnOffset("0"),
                        new RowId("20"),
                        new RowOffset("0")
                    ),
                    new GraphicFrame(
                        new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualGraphicFrameProperties(
                            new NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "Chart 1" },
                            new NonVisualGraphicFrameDrawingProperties()
                        ),
                        new DocumentFormat.OpenXml.Drawing.Spreadsheet.Transform(
                            new A.Offset() { X = 0, Y = 0 },
                            new A.Extents() { Cx = 0, Cy = 0 }
                        ),
                        new A.Graphic(
                            new A.GraphicData(
                                new C.ChartReference() { Id = chartRelId }
                            ) { Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart" }
                        )
                    ),
                    new ClientData()
                );

                drawingsPart.WorksheetDrawing.Append(twoCell);
                drawingsPart.WorksheetDrawing.Save();
            }

            workbookStream.Position = 0;
        }

        private static void GenerateChartPartContent(ChartPart chartPart, string sheetName, string categoryRangeAddress, string valuesRangeAddress, string chartTitle)
        {
            var chartSpace = new C.ChartSpace();
            chartSpace.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
            chartSpace.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            chartSpace.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

            var chart = new C.Chart();

            // Title
            if (!string.IsNullOrEmpty(chartTitle))
            {
                chart.Append(new C.Title(
                    new C.ChartText(new A.TextBody(
                        new A.BodyProperties(),
                        new A.ListStyle(),
                        new A.Paragraph(new A.Run(new A.Text { Text = chartTitle }))
                    ))
                ));
            }

            var plotArea = new C.PlotArea();
            plotArea.Append(new C.Layout());

            // Create a BarChart (clustered column)
            var barChart = new C.BarChart(
                new C.BarDirection() { Val = C.BarDirectionValues.Column },
                new C.BarGrouping() { Val = C.BarGroupingValues.Clustered },
                new C.VaryColors() { Val = false }
            );

            // Add a single series
            var ser = new C.BarChartSeries(
                new C.Index() { Val = (UInt32Value)0U },
                new C.Order() { Val = (UInt32Value)0U },
                new C.SeriesText(new C.StringReference(new C.Formula($"'{sheetName}'!${valuesRangeAddress.Split(':')[0]}")))
            );

            // Category (x axis) - string reference to sheet range
            var cat = new C.CategoryAxisData(
                new C.StringReference(
                    new C.Formula($"'{sheetName}'!{EnsureDollarForAddress(categoryRangeAddress)}")
                )
            );

            // Values - number reference to sheet range
            var val = new C.Values(
                new C.NumberReference(
                    new C.Formula($"'{sheetName}'!{EnsureDollarForAddress(valuesRangeAddress)}")
                )
            );

            ser.Append(cat);
            ser.Append(val);

            barChart.Append(ser);

            // Append axes
            plotArea.Append(barChart);

            // Category Axis (x)
            var catAx = new C.CategoryAxis(
                new C.AxisId() { Val = 48650112U },
                new C.Scaling(new C.Orientation() { Val = C.OrientationValues.MinMax }),
                new C.Delete() { Val = false },
                new C.AxisPosition() { Val = C.AxisPositionValues.Bottom },
                new C.TickLabelPosition() { Val = C.TickLabelPositionValues.NextTo },
                new C.CrossingAxis() { Val = 48672768U },
                new C.Crosses() { Val = C.CrossesValues.AutoZero },
                new C.AutoLabeled() { Val = true },
                new C.LabelAlignment() { Val = C.LabelAlignmentValues.Center },
                new C.LabelOffset() { Val = (UInt16Value)100U }
            );

            // Value Axis (y)
            var valAx = new C.ValueAxis(
                new C.AxisId() { Val = 48672768U },
                new C.Scaling(new C.Orientation() { Val = C.OrientationValues.MinMax }),
                new C.Delete() { Val = false },
                new C.AxisPosition() { Val = C.AxisPositionValues.Left },
                new C.MajorGridlines(),
                new C.NumberingFormat() { FormatCode = "General", SourceLinked = true },
                new C.TickLabelPosition() { Val = C.TickLabelPositionValues.NextTo },
                new C.CrossingAxis() { Val = 48650112U },
                new C.Crosses() { Val = C.CrossesValues.AutoZero },
                new C.CrossBetween() { Val = C.CrossBetweenValues.Between }
            );

            plotArea.Append(catAx);
            plotArea.Append(valAx);

            chart.Append(plotArea);
            chart.Append(new C.PlotVisibleOnly() { Val = true });
            chart.Append(new C.DisplayBlanksAs() { Val = C.DisplayBlanksAsValues.Gap });
            chart.Append(new C.ShowDataLabelsOverMaximum() { Val = false });

            chartSpace.Append(chart);
            chartPart.ChartSpace = chartSpace;
            chartPart.ChartSpace.Save();
        }

        // Ensures addresses are absolute ($A$1:$A$5); if user provided already absolute, leave it
        private static string EnsureDollarForAddress(string address)
        {
            string FixCell(string c)
            {
                int i = 0;
                while (i < c.Length && !char.IsDigit(c[i])) i++;
                var col = c.Substring(0, i);
                var row = c.Substring(i);
                if (!col.StartsWith("$")) col = "$" + col;
                if (!row.StartsWith("$")) row = "$" + row;
                return col + row;
            }

            if (address.Contains(":"))
            {
                var parts = address.Split(':');
                return $"{FixCell(parts[0])}:{FixCell(parts[1])}";
            }
            return FixCell(address);
        }
    }
}