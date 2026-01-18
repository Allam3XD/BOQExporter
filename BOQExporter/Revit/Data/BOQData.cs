using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Media;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.DB;
//using BOQExporter.Excel.Utils;
using BOQExporter.Excel.Utits;
//using BOQExporter.Models;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using static BOQExporter.Revit.Utils.RvtUtils.CategoryUtils;

namespace BOQExporter.Revit.Modeles
{
    internal static class BOQData
    {
        /// <summary>
        /// this represent the Revit application
        /// </summary>
        public static Application App { get; set; }
        /// <summary>
        /// this represent the active Revit document
        /// </summary>
        public static Document Doc { get; set; }
        /// <summary>
        /// this represent the list of all extracted BOQ items
        /// </summary>
        public static List<BOQItem> Items { get; set; } = new List<BOQItem>();
        /// <summary>
        /// this represent the folder that contains all Revit projects , selected by the user.
        /// </summary>
        public static DirectoryInfo RvtProjctDir { get; set; }
        /// <summary>
        /// this represent the folder that the user wants to save the extracted data in.
        /// </summary>
        public static DirectoryInfo SavingDir { get; set; }
        /// <summary>
        /// this represent the file that the extracted data will be saved in.
        /// </summary>
        public static FileInfo ReportFile { get; set; }

        public static void Initialize()
        {
            Items.Clear();
        }

        public static void ExportToExcel(string projectName,
                                 string clientName,
                                 string location,
                                 string issueDate)
        {
            string safeProjectName = string.IsNullOrWhiteSpace(projectName) ? "Project" : projectName.Replace(" ", "_");
            string now = $"{DateTime.Now:yyyy-MM-dd_HH-mm-ss} {(DateTime.Now.Hour >= 12 ? "PM" : "AM")}";
            string fileName = $"{safeProjectName}_BOQ Export_{now}.xlsx";
            ReportFile = new FileInfo(Path.Combine(SavingDir.FullName, fileName));

            using (ExcelPackage ep = new ExcelPackage())
            {
                ExcelWorksheet sheet = ep.Workbook.Worksheets.Add("BOQ Data");

                // 1. Build skeleton (project info + headers)
                XlUtils.ExportBOQ(sheet,
                                  projectName,
                                  clientName,
                                  location,
                                  issueDate,
                                  SelectionMemory.LogoPath);

                // 2. Main BOQ table
                int startRow = 7;
                var groupedMainItems = Items
                    .GroupBy(it => new { it.ItemCode, it.Description, it.Unit, it.Category })
                    .Select((g, index) =>
                    {
                        // Try to extract UnitRate from the first element in the group
                        double unitRate = 0;
                        var firstItem = g.FirstOrDefault();
                        if (firstItem != null && firstItem.Element != null)
                        {
                            Parameter rateParam = firstItem.Element.LookupParameter("UnitRate"); // adjust parameter name if different
                            if (rateParam != null && rateParam.StorageType == StorageType.Double)
                            {
                                unitRate = rateParam.AsDouble();
                            }
                        }

                        return new BOQItem(
                            itemNo: index + 1,
                            category: g.Key.Category,
                            itemCode: g.Key.ItemCode,
                            description: $"{g.Key.Category}:   {g.Key.Description}",
                            quantity: g.Sum(x => x.Quantity),
                            unit: g.Key.Unit,
                            unitRate: unitRate,
                            remarks: string.Empty,
                            element: firstItem?.Element

                        );
                    }).ToList();

                for (int i = 0; i < groupedMainItems.Count; i++)
                {
                    int row = startRow + i;
                    BOQItem item = groupedMainItems[i];
                    if (item.ItemNo == 0) item.ItemNo = i + 1;

                    object[] arr = item.ToArray();
                    for (int col = 0; col < arr.Length; col++)
                        sheet.Cells[row, col + 1].Value = arr[col];
                }

                // 3. Subtotal row
                int subtotalRow = startRow + groupedMainItems.Count;
                string MainColumnName = XlUtils.GetColumnName(7);

                sheet.Cells[subtotalRow, 6].Value = "Subtotal:";
                sheet.Cells[subtotalRow, 7].Formula = $"SUM({MainColumnName}{startRow}:{MainColumnName}{subtotalRow - 1})";

                // Apply currency formatting to Unit Rate (F) and Total Cost (G) columns
                sheet.Cells[$"F{startRow}:F{subtotalRow - 1}"].Style.Numberformat.Format = "#,##0.00 [$EGP]";
                sheet.Cells[$"G{startRow}:G{subtotalRow}"].Style.Numberformat.Format = "#,##0.00 [$EGP]"; 

                // Apply subtotal styling (no redundant bold here)
                XlUtils.ApplyTotalRowStyle(sheet, $"A{subtotalRow}", $"H{subtotalRow}");

                int lastRow = subtotalRow;

                // Apply table styling
                XlUtils.ApplyTableBorders(sheet, "A6", $"H{lastRow}");
                XlUtils.ApplyHeaderStyle(sheet, "A6", "H6");
                XlUtils.ApplyAlternatingRowShading(sheet, startRow, subtotalRow - 1, 1, 8); // ✅ exclude subtotal

                // Column alignment for main BOQ table
                sheet.Cells["A6:H6"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center; // header row

                sheet.Cells[$"A{startRow}:A{subtotalRow - 1}"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center; // Item No.
                sheet.Cells[$"B{startRow}:B{subtotalRow - 1}"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center; // Item Code
                sheet.Cells[$"C{startRow}:C{subtotalRow - 1}"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;   // Description
                sheet.Cells[$"D{startRow}:H{subtotalRow - 1}"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center; // Numeric values

                //making the remarks column wide enough

                // 4. Traceability table
                int traceStrartRow = subtotalRow + 3;
                XlUtils.CreateTraceabilityHeader(sheet, traceStrartRow);

                // Group items
                var groupedItems = Items
                    .GroupBy(it => new { it.Category, it.ItemCode, it.Description, it.Unit })
                    .Select(g => new BOQItem(
                        category: g.Key.Category,
                        itemCode: g.Key.ItemCode,
                        description: g.Key.Description, // ✅ no prefix now
                        quantity: g.Sum(x => x.Quantity),
                        unit: g.Key.Unit
                    )
                    { UsageCount = g.Count() }).ToList();

                // Fill rows
                int traceRow = traceStrartRow + 2; // header row is traceStrartRow+1
                foreach (var gItem in groupedItems)
                { 
                    object[] arr = gItem.ToTraceArray(); // now includes Category
                    for (int col = 0; col < arr.Length; col++)
                        sheet.Cells[traceRow, col + 1].Value = arr[col];
                    traceRow++;
                }

                // Totals row
                int totalItemRow = traceRow;
                sheet.Cells[totalItemRow, 5].Value = "Total Items:";
                sheet.Cells[totalItemRow, 6].Formula = $"SUM(F{traceStrartRow + 2}:F{totalItemRow - 1})";
                XlUtils.ApplyTotalRowStyle(sheet, $"A{totalItemRow}", $"F{totalItemRow}");

                // Styling
                XlUtils.ApplyTableBorders(sheet, $"A{traceStrartRow + 1}", $"F{totalItemRow}");
                XlUtils.ApplyHeaderStyle(sheet, $"A{traceStrartRow + 1}", $"F{traceStrartRow + 1}");
                XlUtils.ApplyAlternatingRowShading(sheet, traceStrartRow + 2, totalItemRow - 1, 1, 6);
                // Alignment
                sheet.Cells[$"A{traceStrartRow + 1}:F{traceStrartRow + 1}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; // header
                sheet.Cells[$"A{traceStrartRow + 2}:A{totalItemRow - 1}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; // Category
                sheet.Cells[$"B{traceStrartRow + 2}:B{totalItemRow - 1}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; // Item Code
                sheet.Cells[$"C{traceStrartRow + 2}:C{totalItemRow - 1}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;   // Description
                sheet.Cells[$"D{traceStrartRow + 2}:F{totalItemRow - 1}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; // Numeric values

                // 5. Category summary chart
                var categoryCounts = Items
                    .GroupBy(it => it.Category)
                    .Select(g => new { Category = g.Key, Count = g.Count() })
                    .ToList();

                XlUtils.AddPieChart(sheet, startRow - 1, 10, categoryCounts, "Category Distribution");

                // 6. Auto-fit columns
                sheet.Cells[sheet.Dimension.Address].AutoFitColumns();
                sheet.Column(8).Width = 30;

                // 7. Save file
                ep.SaveAs(ReportFile);
                try
                {
                    System.Diagnostics.Process.Start(ReportFile.FullName);
                }
                catch { /* optional: handle non-Windows environments */ }
            }
        }

    }
}
