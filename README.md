# ClosedXML.Charts (extension library)

This is a small sample .NET Framework class library that demonstrates adding a chart into a ClosedXML-created workbook by post-processing the .xlsx package with the Open XML SDK.

Usage example:

1. Create and populate a ClosedXML workbook as usual:

```csharp
using ClosedXML.Excel;
using ClosedXML.Charts;
using System.IO;

var wb = new XLWorkbook();
var ws = wb.Worksheets.Add("Sheet1");
ws.Cell("A1").Value = "Category";
ws.Cell("B1").Value = "Value";
ws.Cell("A2").Value = "Alpha";
ws.Cell("B2").Value = 10;
ws.Cell("A3").Value = "Beta";
ws.Cell("B3").Value = 30;
ws.Cell("A4").Value = "Gamma";
ws.Cell("B4").Value = 20;
ws.Cell("A5").Value = "Delta";
ws.Cell("B5").Value = 40;

// Save workbook and inject chart
using var ms = wb.SaveWithBarChart("Sheet1", "A2:A5", "B2:B5", "My Values Chart");
using var fs = File.Create("with-chart.xlsx");
ms.CopyTo(fs);
```

2. The helper will add a clustered column chart that references the ranges provided.

Notes and limitations:
- This is a starting implementation. The chart anchor coordinates are simple fixed values; you may need to compute better positions based on cell sizes if you need precise placement.
- The code inserts a clustered column chart only. Other chart types (line, pie, stacked, scatter) can be created by changing the chart XML (see Open XML Chart schema).
- The helper assumes a simple address format like "A2:A5" for ranges. Sheet names with special characters are quoted.
- Make sure you use compatible versions of ClosedXML and DocumentFormat.OpenXml.