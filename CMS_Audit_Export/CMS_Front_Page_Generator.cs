using System.Data;
using System.Drawing;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace CMS_Front_Page_Generator
{
    class CMS_Front_Page_Generator
    {
        public static void Add_Frontpage_row(ExcelWorksheet ws, DataRow control, int row)
        {
#pragma warning disable CS8600 // Converting null literal or possible null value to non-nullable type.
            string page = control["Control"].ToString();
            ws.Cells[row, 1].Value = row - 20;
            ws.Cells[row, 2].Hyperlink = new Uri("#'" + control["Control"].ToString() + "'!A5", UriKind.Relative);
            ws.Cells[row, 2].Value = control["Control"].ToString();
            ws.Cells[row, 3].Value = control["Name"];
#pragma warning restore CS8600 // Converting null literal or possible null value to non-nullable type.
            ws.Cells[row, 4].Formula = "=IF('" + page + "'!B$25=\"\",\"\",'" + page + "'!B$25)";
            ws.Cells[row, 5].Formula = "=IF('" + page + "'!B$32=\"\",\"\",'" + page + "'!B$32)";
            ws.Cells[row, 6].Formula = "=IF('" + page + "'!B$33=\"\",\"\",'" + page + "'!B$33)";
            ws.Cells[row, 6].Style.Numberformat.Format = "MM/DD/YYYY";
            ws.Cells[row, 7].Formula = "=IF('" + page + "'!B$34=\"\",\"\",'" + page + "'!B$34)";
            ws.Cells[row, 7].Style.Numberformat.Format = "MM/DD/YYYY";
            ws.Cells[row, 8].Formula = "=IF('" + page + "'!B$39=\"\",\"\",'" + page + "'!B$39)";
            ws.Cells[row, 9].Formula = "=IF('" + page + "'!B$45=\"\",\"\",'" + page + "'!B$45)";
            ws.Cells[row, 10].Formula = "=IF('" + page + "'!B$46=\"\",\"\",'" + page + "'!B$46)";
        }

        public static void Add_Frontpage_Header(ExcelWorksheet ws, int row, DataTable PoC, string title, string version)
        {
            int cells = row - 20;
            ws.Column(1).Width = 8.75;
            ws.Column(2).Width = 8.75;
            ws.Column(3).Width = 60.5;
            ws.Column(4).Width = 13.75;
            ws.Column(5).Width = 9.5;
            ws.Column(6).Width = 11.25;
            ws.Column(7).Width = 11.25;
            ws.Column(8).Width = 8.25;
            ws.Column(9).Width = 7.25;
            ws.Column(10).Width = 11.5;

            Add_Frontpage_Title(ws, title);
            Add_Frontpage_PoC(ws, PoC, version);
            Add_Frontpage_Instructions(ws);
            Add_Frontpage_Controls(ws);
            Add_Frontpage_Totals(ws, cells, row);
        }

        private static void Add_Frontpage_Title(ExcelWorksheet ws, string title)
        {
            ws.Cells["A1:J1"].Merge = true;
            ws.Cells["A1:J1"].Value = title;
            ws.Cells["A1:J1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Cells["A1:J1"].Style.Font.Bold = true;
            ws.Cells["A1:J1"].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
            ws.Cells["A1:J1"].Style.Border.Right.Style = ExcelBorderStyle.Medium;

        }

        private static void Add_Frontpage_PoC(ExcelWorksheet ws, DataTable PoC, string version)
        {
            ws.Cells["A4:A9"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["A4:A9"].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#808080"));
            ws.Cells["A4:A9"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
            ws.Cells[4, 1].Value = "PoC:";
            ws.Cells[6, 1].Value = "Platform:";
            ws.Cells[7, 1].Value = "Function:";
            ws.Cells[9, 1].Value = "Date:";

            foreach (DataRow i in PoC.Rows)
            {
                ws.Cells[4, 2].Value = i["TEAM_ID"].ToString();
                ws.Cells[7, 2].Value = i["SYSTEM_ID"].ToString();
                ws.Cells[6, 2].Value = i["PLATFORM"].ToString();

                // Add secret cells for tracking
                ws.Cells["F4"].Value = i["AGENCY_ID"];
                try
                {
                    ws.Cells["F5"].Value = Int32.Parse(version);
                }
                catch
                {
                    ws.Cells["F5"].Value = version;
                }

                // Force validation
                var AgencyID = ws.DataValidations.AddListValidation("F4");
                AgencyID.Formula.Values.Add(i["AGENCY_ID"].ToString());
                var Version = ws.DataValidations.AddListValidation("F5");
                Version.Formula.Values.Add(version.ToString());
            }
            ws.Cells[9, 2].Value = DateTime.Today.ToShortDateString();

            // Make the tracking data hidden
            ws.Cells["F4"].Style.Font.Color.SetColor(Color.White);
            ws.Cells["F5"].Style.Font.Color.SetColor(Color.White);

        }

        private static void Add_Frontpage_Instructions(ExcelWorksheet ws)
        {
            ws.Cells["A11:G11"].Merge = true;
            ws.Cells["A11:G11"].Style.Font.Bold = true;
            ws.Cells["A11:G11"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["A11:G11"].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#808080"));
            ws.Cells["A11:G11"].Value = "Instructions";
            ws.Cells["A12:G12"].Merge = true;
            ws.Cells["A12:G12"].Value = "Use this worksheet to navigate the workbook and to track your effort.";
            ws.Cells["A13:G13"].Merge = true;
            ws.Cells["A13:G13"].Value = "    Navigate to an individual worksheet by clicking on the \"Control Id\" in column B of the list below.";
            ws.Cells["A14:G14"].Merge = true;
            ws.Cells["A14:G14"].Value = "    Track your effort and the Audit results on this worksheet:";
            ws.Cells["A15:G15"].Merge = true;
            ws.Cells["A15:G15"].Value = "        Pre-Audit and Audit columns in the list below are populated from information entered on the individual control worksheets.";
            ws.Cells["A16:G16"].Merge = true;
            ws.Cells["A16:G16"].Value = "        Statistics on progress are displayed above and to the right of this worksheet.";
            ws.Cells["A12:G16"].Style.Font.Color.SetColor(Color.Red);
        }

        private static void Add_Frontpage_Controls(ExcelWorksheet ws)
        {
            ws.Cells["D18:G18"].Merge = true;
            ws.Cells["D18:G18"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["D18:G18"].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#808080"));
            ws.Cells["D18:G18"].Value = "Pre-Audit";
            ws.Cells["D18:G18"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            Set_all_borders(ws, "D18:G18");

            ws.Cells["H18:J18"].Merge = true;
            ws.Cells["H18:J18"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["H18:J18"].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#808080"));
            ws.Cells["H18:J18"].Value = "Audit Results";
            ws.Cells["H18:J18"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            Set_all_borders(ws, "H18:J18");

            ws.Cells["A19:E19"].Merge = true;
            ws.Cells["A19:E19"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["A19:E19"].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#808080"));
            ws.Cells["A19:E19"].Value = "Summary List of Controls";
            Set_all_borders(ws, "A19:E19");

            ws.Cells["F19:G19"].Merge = true;
            ws.Cells["F19:G19"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["F19:G19"].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#C0C0C0"));
            ws.Cells["F19:G19"].Value = "Date Completed";
            ws.Cells["F19:G19"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            Set_all_borders(ws, "F19:G19");

            ws.Cells["H19"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["H19"].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#808080"));
            ws.Cells["H19"].Value = "Previous";
            ws.Cells["H19"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            Set_all_borders(ws, "H19");

            ws.Cells["I19:J19"].Merge = true;
            ws.Cells["I19:J19"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["I19:J19"].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#808080"));
            ws.Cells["I19:J19"].Value = "Current";
            ws.Cells["I19:J19"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            Set_all_borders(ws, "I19:J19");

            // Insert Headers
            ws.Cells["A20:J20"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["A20:J20"].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#C0C0C0"));
            ws.Cells[20, 1].Value = "Number";
            Set_all_borders(ws, "A20");
            ws.Cells[20, 2].Value = "Control";
            Set_all_borders(ws, "B20");
            ws.Cells[20, 3].Value = "Name";
            Set_all_borders(ws, "C20");
            ws.Cells[20, 4].Value = "Applicable";
            Set_all_borders(ws, "D20");
            ws.Cells[20, 5].Value = "Status";
            Set_all_borders(ws, "E20");
            ws.Cells[20, 6].Value = "Description";
            Set_all_borders(ws, "F20");
            ws.Cells[20, 7].Value = "Evidence";
            Set_all_borders(ws, "G20");
            ws.Cells[20, 8].Value = "Results";
            Set_all_borders(ws, "H20");
            ws.Cells[20, 9].Value = "Results";
            Set_all_borders(ws, "I20");
            ws.Cells[20, 10].Value = "Criticality";
            Set_all_borders(ws, "J20");
        }

        private static void Add_Frontpage_Totals(ExcelWorksheet ws, int cells, int row)
        {
            ws.Cells["G4:G9"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["G4:G7"].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#808080"));
            ws.Cells["G8:G9"].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#808080"));
            ws.Cells["G4:G9"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
            ws.Cells["G4:I4"].Merge = true;
            ws.Cells[4, 7].Value = "Total Controls:";
            ws.Cells[4, 10].Value = cells;
            ws.Cells["G5:I5"].Merge = true;
            ws.Cells[5, 7].Value = "Pre-Audit Complete:";
            ws.Cells[5, 10].Formula = "=COUNTIF(E21:E" + row.ToString() + ",\"Complete\")";
            ws.Cells["G6:I6"].Merge = true;
            ws.Cells[6, 7].Value = "Pre-Audit % Complete";
            ws.Cells[6, 10].Formula = "=J5/J4";
            ws.Cells[6, 10].Style.Numberformat.Format = "0%";
            ws.Cells["G8:I8"].Merge = true;
            ws.Cells[8, 7].Value = "Audit - Results:";
            ws.Cells[8, 10].Formula = "=J4-COUNTBLANK(I21:I" + row.ToString() + ")";
            ws.Cells["G9:I9"].Merge = true;
            ws.Cells[9, 7].Value = "Audit % Results:";
            ws.Cells[9, 10].Formula = "=J8/J4";
            ws.Cells[9, 10].Style.Numberformat.Format = "0%";
        }

        private static void Set_all_borders(ExcelWorksheet ws, string block)
        {
            ws.Cells[block].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
            ws.Cells[block].Style.Border.Left.Style = ExcelBorderStyle.Medium;
            ws.Cells[block].Style.Border.Right.Style = ExcelBorderStyle.Medium;
            ws.Cells[block].Style.Border.Top.Style = ExcelBorderStyle.Medium;
        }
    }
}
