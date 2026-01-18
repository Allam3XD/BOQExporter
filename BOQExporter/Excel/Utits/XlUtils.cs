using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Media.Animation;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace BOQExporter.Excel.Utits
{
    internal class XlUtils
    {
        /// <summary>
        /// to chose the coloer
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="red"></param>
        /// <param name="green"></param>
        /// <param name="blue"></param>
        public static void SetCellColor(ExcelRange cell, int red, int green, int blue)
        {
            cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
            cell.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(red, green, blue));
        }



        /// <summary>
        /// creating a bold cell
        /// </summary>
        /// <param name="cell"></param>
        public static void SetBold(ExcelRange cell)
        {
            cell.Style.Font.Bold = true;
        }



        /// <summary>
        /// to set the currency format
        /// </summary>
        /// <param name="cell"></param>
        public static void SetCurrencyFormat(ExcelRange cell)
        {
            cell.Style.Numberformat.Format = "#,##0.00";
        }




        /// <summary>
        /// creating a header
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="fromRow"></param>
        /// <param name="fromCol"></param>
        /// <param name="headerFields"></param>
        public static void CreateHeader(ExcelWorksheet sheet,
                                        int fromRow,
                                        int fromCol,
                                        params string[] headerFields)
        {
            for (int i = 0; i < headerFields.Length; i++)
                sheet.Cells[fromRow, fromCol + i].Value = headerFields[i];

            ExcelRange headerCells = sheet.Cells[fromRow, fromCol, fromRow, fromCol + headerFields.Length - 1];
            SetCellColor(headerCells, 191, 191, 191);  // Gray
            SetBold(headerCells);
        }




        /// <summary>
        /// this method sets up the basic structure of the BOQ Excel sheet,
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="projectName"></param>
        /// <param name="clientName"></param>
        /// <param name="location"></param>
        /// <param name="issueDate"></param>
        /// <param name="logoPath"></param>
        public static void ExportBOQ(ExcelWorksheet sheet,
                                    string projectName,
                                    string clientName,
                                    string location,
                                    string issueDate,
                                    string logoPath)
        {
            // --- Project Info Section ---
            // Project Name row
            sheet.Cells["A1:C1"].Merge = true;
            sheet.Cells["A1"].Value = $"Project Name: {projectName}";
            sheet.Cells["A1"].Style.Font.Bold = true;
            sheet.Cells["A1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheet.Cells["A1"].Style.Fill.BackgroundColor.SetColor(Color.LightGray);

            // Client row
            sheet.Cells["A2:C2"].Merge = true;
            sheet.Cells["A2"].Value = $"Client: {clientName}";
            sheet.Cells["A2"].Style.Font.Bold = true;
            sheet.Cells["A2"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheet.Cells["A2"].Style.Fill.BackgroundColor.SetColor(Color.LightGray);

            // Location row
            sheet.Cells["A3:C3"].Merge = true;
            sheet.Cells["A3"].Value = $"Location: {location}";
            sheet.Cells["A3"].Style.Font.Bold = true;
            sheet.Cells["A3"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheet.Cells["A3"].Style.Fill.BackgroundColor.SetColor(Color.LightGray);

            // Date row
            sheet.Cells["A4:C4"].Merge = true;
            sheet.Cells["A4"].Value = $"Date: {issueDate}";
            sheet.Cells["A4"].Style.Font.Bold = true;
            sheet.Cells["A4"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheet.Cells["A4"].Style.Fill.BackgroundColor.SetColor(Color.LightGray);


            //add logo if provided
            if (!string.IsNullOrEmpty(logoPath) && File.Exists(logoPath))
            {
                var picture = sheet.Drawings.AddPicture("Logo", Image.FromFile(logoPath));

                // merge D–E for logo space
                sheet.Cells["D1:E4"].Merge = true;
                

                // position logo in merged block
                picture.SetPosition(0, 5, 3, 5); // row, pixelOffsetY, column, pixelOffsetX
                picture.SetSize(150, 75);        // adjust size
            }

            //the borders  for the metadata section 
            // Apply thin outer border around the header block (A1:C4)
            var headerRange = sheet.Cells["A1:C4"];
            headerRange.Style.Border.Top.Style = ExcelBorderStyle.Thin;
            headerRange.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            headerRange.Style.Border.Left.Style = ExcelBorderStyle.Thin;
            headerRange.Style.Border.Right.Style =  ExcelBorderStyle.Thin;



            // Add sheet title
            sheet.Cells["A5"].Value = "BOQ Report";
            sheet.Cells["A5:H5"].Merge = true;
            sheet.Cells["A5"].Style.Font.Bold = true;
            sheet.Cells["A5"].Style.Font.Size = 14;
            sheet.Cells["A5"].Style.HorizontalAlignment =ExcelHorizontalAlignment.Center;

            // --- BOQ Headers ---
            string[] headers = { 
                                "Item No.",
                                "Item Code",
                                "Description",
                                "Quantity",
                                "Unit",
                                "Unit Rate",
                                "Total Cost",
                                "Remarks" 
                                };

            CreateHeader(sheet, 6, 1, headers);

            // --- Formatting ---
            sheet.Cells[sheet.Dimension.Address].AutoFitColumns();
            //sheet.View.FreezePanes(7, 1);
        }
        /// <summary>
        /// this method creates headers for the traceability section in the Excel sheet.
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="fromRow"></param>
        public static void CreateTraceabilityHeader(ExcelWorksheet sheet, int fromRow)
        {
            // Title row
            sheet.Cells[$"A{fromRow}"].Value = "BOQ Summary Report";
            sheet.Cells[$"A{fromRow}:F{fromRow}"].Merge = true;
            sheet.Cells[$"A{fromRow}"].Style.Font.Bold = true;
            sheet.Cells[$"A{fromRow}"].Style.Font.Size = 14;
            sheet.Cells[$"A{fromRow}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            // Header row (next line)
            string[] headers = {
                                "Item Code",
                                "Category",
                                "Description",
                                "Quantity",
                                "Unit",
                                "Usage Count"
                               };

            CreateHeader(sheet, fromRow + 1, 1, headers);
        }


        /// <summary>
        /// adds a pie chart to the given worksheet based on the provided summary data.
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="fromRow"></param>
        /// <param name="fromCol"></param>
        /// <param name="summaryList"></param>
        /// <param name="chartTitle"></param>
        public static void AddPieChart(ExcelWorksheet sheet, int fromRow, int fromCol,
                       IEnumerable<dynamic> summaryList, string chartTitle)
        {
            // title row
            var titleRow = fromRow-1;
            sheet.Cells[$"J{titleRow}"].Value = "Category Analysis";
            sheet.Cells[$"J{titleRow}:L{titleRow}"].Merge = true;
            sheet.Cells[$"J{titleRow}"].Style.Font.Bold = true;
            sheet.Cells[$"J{titleRow}"].Style.Font.Size = 14;
            sheet.Cells[$"J{titleRow}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            // Headers
            sheet.Cells[fromRow, fromCol].Value = "Label";
            sheet.Cells[fromRow, fromCol + 1].Value = "Value";
            sheet.Cells[fromRow, fromCol + 2].Value = "Percentage";

            ExcelRange headerCells = sheet.Cells[fromRow, fromCol, fromRow, fromCol + 2];
            SetCellColor(headerCells, 0, 255, 0);  // green
            SetBold(headerCells);

            // Compute total (sum of values)
            // Safely extract numeric value from dynamic: prefer Count, fallback to Value
            double total = summaryList.Sum(item =>
            {
                object raw = item.Count ?? item.Value;
                return raw is double d ? d
                     : raw is int i ? (double)i
                     : raw is long l ? (double)l
                     : 0.0;
            });

            // Fill rows
            int row = fromRow + 1;
            foreach (var item in summaryList)
            {
                string label = item.Category ?? item.Label;
                object raw = item.Count ?? item.Value;

                double value = raw is double d ? d
                             : raw is int i ? (double)i
                             : raw is long l ? (double)l
                             : 0.0;

                sheet.Cells[row, fromCol].Value = label;
                sheet.Cells[row, fromCol + 1].Value = value;

                // Percentage = value / total (guard against zero)
                double pct = total > 0 ? value / total : 0.0;
                sheet.Cells[row, fromCol + 2].Value = pct;

                row++;
            }

            // Format percentage column
            sheet.Cells[fromRow + 1, fromCol + 2, row - 1, fromCol + 2].Style.Numberformat.Format = "0%";

            // Borders & shading (cover all 3 columns)
            ApplyTableBorders(sheet, $"{GetColumnName(fromCol)}{fromRow}", $"{GetColumnName(fromCol + 2)}{row - 1}");
            ApplyAlternatingRowShading(sheet, fromRow + 1, row - 1, fromCol, fromCol + 2);

            // Create pie chart
            var pieChart = sheet.Drawings.AddChart(chartTitle.Replace(" ", "") + "PieChart",
                                                   OfficeOpenXml.Drawing.Chart.eChartType.Pie3D);

            pieChart.Title.Text = chartTitle;
            pieChart.SetPosition(fromRow - 1, 0, fromCol + 3, 0);
            pieChart.SetSize(600, 400);

            // Series (Values vs Labels)
            var series = pieChart.Series.Add(
                sheet.Cells[fromRow + 1, fromCol + 1, row - 1, fromCol + 1], // Y values
                sheet.Cells[fromRow + 1, fromCol, row - 1, fromCol]      // X labels
            );
            series.Header = chartTitle;

            // Data labels
            var pieSeries = series as OfficeOpenXml.Drawing.Chart.ExcelPieChartSerie;
            if (pieSeries != null)
            {
                pieSeries.DataLabel.ShowPercent = false;
                pieSeries.DataLabel.ShowCategory = true;
                pieSeries.DataLabel.ShowValue = false;
                pieSeries.DataLabel.Separator = "\n";
            }
        }





        /// <summary>
        /// converts a column number to its corresponding Excel column name.
        /// </summary>
        /// <param name="columnNumber"></param>
        /// <returns></returns>
        public static string GetColumnName(int columnNumber)
        {
            string columnName = "";
            while (columnNumber > 0)
            {
                int remainder = (columnNumber - 1) % 26;
                columnName = (char)(remainder + 'A') + columnName;
                columnNumber = (columnNumber - 1) / 26;
            }
            return columnName;
        }
        /// <summary>
        /// this method applies thin borders around all cells in the specified table range.
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="firstCell"></param>
        /// <param name="lastCell"></param>
        public static void ApplyTableBorders(ExcelWorksheet sheet, string firstCell, string lastCell)
        {
            var tableRange = sheet.Cells[$"{firstCell}:{lastCell}"];

            // Apply thin borders around all cells in the table
            tableRange.Style.Border.Top.Style = ExcelBorderStyle.Thin;
            tableRange.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            tableRange.Style.Border.Left.Style = ExcelBorderStyle.Thin;
            tableRange.Style.Border.Right.Style = ExcelBorderStyle.Thin;
        }
        /// <summary>
        /// this method applies a header style to the specified range in the worksheet.
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="firstCell"></param>
        /// <param name="lastCell"></param>
        public static void ApplyHeaderStyle(ExcelWorksheet sheet,
                                            string firstCell,
                                            string lastCell)
        {
            var headerRange = sheet.Cells[$"{firstCell}:{lastCell}"];

            headerRange.Style.Font.Bold = true;
            headerRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
            headerRange.Style.Fill.BackgroundColor.SetColor(Color.LightGray);
            headerRange.Style.Border.Bottom.Style = ExcelBorderStyle.Thick;
           
        }
        /// <summary>
        /// this method applies alternating row shading to enhance readability in the specified range.
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="startRow"></param>
        /// <param name="endRow"></param>
        /// <param name="startCol"></param>
        /// <param name="endCol"></param>
        public static void ApplyAlternatingRowShading(ExcelWorksheet sheet,
                                                    int startRow,
                                                    int endRow,
                                                    int startCol,
                                                    int endCol)
        {
            for (int row = startRow; row <= endRow; row++)
            {
                if ((row - startRow) % 2 == 0) // even rows relative to start
                {
                    var rowRange = sheet.Cells[row, startCol, row, endCol];
                    rowRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    rowRange.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(240, 240, 240)); // light gray
                }
            }
        }


        /// <summary>
        /// this method applies a distinct style to the total row in the worksheet.
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="firstCell"></param>
        /// <param name="lastCell"></param>
        public static void ApplyTotalRowStyle(ExcelWorksheet sheet, string firstCell, string lastCell)
        {
            var totalRange = sheet.Cells[$"{firstCell}:{lastCell}"];

            totalRange.Style.Font.Bold = true;
            totalRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
            totalRange.Style.Fill.BackgroundColor.SetColor(Color.LightYellow);
            totalRange.Style.Border.Top.Style = ExcelBorderStyle.Thick;
        }

    }
}
