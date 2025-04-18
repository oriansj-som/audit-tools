using common;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Data;
using System.Drawing;

namespace CMS_Controls_Export
{
    public class CMS_Controls_Export
    {
        public SQLiteDatabase? mydatabase;
        public ExcelPackage? Report;
        public FileInfo? output;

        public void Initialize(string Database)
        {
            Require.True(File.Exists(Database), "Specified database doesn't exist");

            // Create our new temp excel file
            Report = new ExcelPackage();

            // Connect to Database
            mydatabase = new SQLiteDatabase(Database);
        }
        public int Initialize(string Database, string Audit_Name)
        {
            Initialize(Database);

            try
            {
                string sql = string.Format("SELECT AUDIT_TYPE FROM AUDIT_TYPES WHERE UPPER(NAME) = UPPER('{0}');", Audit_Name);
                Require.False(null == mydatabase, "Unable to connect to database");
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                string hold = mydatabase.ExecuteScalar(sql);
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                int r = Convert.ToInt32(hold);
                return r;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.WriteLine("unable to find AuditType ID for provided Name");
                Environment.Exit(1);
            }
            // Should never be hit
            throw new Exception("Impossible case: Initialize 1");
        }

        public string Initialize(string Database, int AuditID)
        {
            Initialize(Database);

            try
            {
                string sql = string.Format("SELECT NAME FROM AUDIT_TYPES WHERE AUDIT_TYPE = {0};", AuditID);
                Require.False(null == mydatabase, "Unable to connect to database");
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                string r = mydatabase.ExecuteScalar(sql);
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                return r;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.WriteLine("unable to find AuditType ID for provided Name");
                Environment.Exit(1);
            }
            // Should never be hit
            throw new Exception("Impossible case: Initialize 2");
        }

        public void Generate(FileInfo o, int audit_number, string audit_name)
        {
            if (null == mydatabase) return;
            output = o;
            if (Match.CaseInsensitive("CMS ARC-AMPE", audit_name)) Export_arc_ampe(audit_number, audit_name);
            else if (Match.CaseInsensitive("CMS MARS-E 2.2", audit_name)) Export_CMS(audit_number, audit_name);
            else if (Match.CaseInsensitive("CMS MARS-E 2.0", audit_name)) Export_CMS(audit_number, audit_name);
            else
            {
                Console.WriteLine(string.Format("Unknown Control Type", audit_name));
                Environment.Exit(1);
            }
        }

        private void Export_CMS(int audit_number, string audit_name)
        {
            string sql = string.Format("SELECT CONTROL_ID, FAMILY_ID, FAMILY_NAME, CONTROL, TYPE, CONTROLS.NAME, DESCRIPTION, IMPLEMENTATION_STANDARDS, GUIDANCE, RELATED_CONTROL_REQUIREMENTS, ASSESSMENT_OBJECTIVE, ASSESSMENT_METHOD, PRIOR_ID "
                                      + "FROM CONTROLS "
                                      + "WHERE AUDIT_TYPE = {0};"
                                      , audit_number);
            DataTable results = mydatabase.GetDataTable(sql);

            foreach (DataRow row in results.Rows)
            {
                Populate_CMS_control(row, audit_name);
            }
        }

        private void Populate_CMS_control(DataRow row, string audit_name)
        {
#pragma warning disable CS8602 // Dereference of a possibly null reference.
            Require.False(null == Report, "Report file improperly initialized");
            Require.False(null == Report.Workbook, "Report file defective");
            Require.False(null == Report.Workbook.Worksheets, "Report sheets not initialized");
#pragma warning restore CS8602 // Dereference of a possibly null reference.

            if (null == row) return;
            if (null == Report) return;
            if (null == Report.Workbook) return;
            if (null == Report.Workbook.Worksheets) return;

            // Name it
#pragma warning disable CS8600 // Converting null literal or possible null value to non-nullable type.
            string Control = row["CONTROL"].ToString();
            Report.Workbook.Worksheets.Add(Control);
            ExcelWorksheet ws = Report.Workbook.Worksheets[Control];

            // fill it
#pragma warning disable CS8604 // Possible null reference argument.
            string Control_ID = row["CONTROL_ID"].ToString();
            Populate_row(ws, 2, "Control ID", Control_ID);
            string Family_ID = row["FAMILY_ID"].ToString();
            Populate_row(ws, 3, "Family ID", Family_ID);
            string Family_Name = row["FAMILY_NAME"].ToString();
            Populate_row(ws, 4, "Family Name", Family_Name);
            // Use it
            Populate_row(ws, 5, "Control", Control);
            string Type = row["TYPE"].ToString();
            Populate_row(ws, 6, "Type", Type);
            string Name = row["NAME"].ToString();
            Populate_row(ws, 7, "Name", Name);
            string Description = row["DESCRIPTION"].ToString();
            Populate_row(ws, 8, "Description", Description);
            string Implementation_Standards = row["IMPLEMENTATION_STANDARDS"].ToString();
            Populate_row(ws, 9, "Implementation Standards", Implementation_Standards);
            string Guidance = row["GUIDANCE"].ToString();
            Populate_row(ws, 10, "Guidance", Guidance);
            string Related_Control_Requirements = row["RELATED_CONTROL_REQUIREMENTS"].ToString();
            Populate_row(ws, 11, "Related Control Requirements", Related_Control_Requirements);
            string Assessment_Objective = row["ASSESSMENT_OBJECTIVE"].ToString();
            Populate_row(ws, 12, "Assessment Objective", Assessment_Objective);
            string Assessment_Method = row["ASSESSMENT_METHOD"].ToString();
            Populate_row(ws, 13, "Assessment Method", Assessment_Method);
            string Prior_ID = row["PRIOR_ID"].ToString();
            Populate_row(ws, 14, "Prior ID", Prior_ID);
#pragma warning restore CS8604 // Possible null reference argument.
#pragma warning restore CS8600 // Converting null literal or possible null value to non-nullable type.


            // Add some style
            ws.Column(1).Width = 40.71;
            ws.Column(2).Width = 139.29;

            // Make the header
            ws.Cells["A1:B1"].Merge = true;
            ws.Cells["A1:B1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Cells["A1:B1"].Style.Font.Bold = true;
            ws.Cells["A1:B1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["A1:B1"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#909090"));
            ws.Cells["A1:B1"].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
            ws.Cells["A1:B1"].Style.Border.Right.Style = ExcelBorderStyle.Medium;
            ws.Cells["A1:B1"].Value = string.Format("{0} Control Information:", audit_name);
        }

        private void Export_arc_ampe(int audit_number, string audit_name)
        {
            string sql = string.Format("SELECT CONTROL_ID, FAMILY_ID, FAMILY_NAME, CONTROL, TYPE, CONTROLS.NAME, DESCRIPTION, AMPE_GUIDANCE, NIST_GUIDANCE, RELATED_CONTROL_REQUIREMENTS, PRIOR_ID "
                                     + "FROM CONTROLS "
                                     + "WHERE AUDIT_TYPE = {0};"
                                     ,audit_number);
            DataTable results = mydatabase.GetDataTable(sql);

            foreach (DataRow row in results.Rows)
            {
                Populate_arc_ampe_control(row, audit_name);
            }
        }

        private void Populate_arc_ampe_control(DataRow row, string audit_name)
        {
#pragma warning disable CS8602 // Dereference of a possibly null reference.
            Require.False(null == Report, "Report file improperly initialized");
            Require.False(null == Report.Workbook, "Report file defective");
            Require.False(null == Report.Workbook.Worksheets, "Report sheets not initialized");
#pragma warning restore CS8602 // Dereference of a possibly null reference.

            if (null == row) return;
            if (null == Report) return;
            if (null == Report.Workbook) return;
            if (null == Report.Workbook.Worksheets) return;

            // Name it
#pragma warning disable CS8600 // Converting null literal or possible null value to non-nullable type.
            string Control = row["CONTROL"].ToString();
            Report.Workbook.Worksheets.Add(Control);
            ExcelWorksheet ws = Report.Workbook.Worksheets[Control];

            // fill it
#pragma warning disable CS8604 // Possible null reference argument.
            string Control_ID = row["CONTROL_ID"].ToString();
            Populate_row(ws, 2, "Control ID", Control_ID);
            string Family_ID = row["FAMILY_ID"].ToString();
            Populate_row(ws, 3, "Family ID", Family_ID);
            string Family_Name = row["FAMILY_NAME"].ToString();
            Populate_row(ws, 4, "Family Name", Family_Name);
            // Use it
            Populate_row(ws, 5, "Control", Control);
            string Type = row["TYPE"].ToString();
            Populate_row(ws, 6, "Type", Type);
            string Name = row["NAME"].ToString();
            Populate_row(ws, 7, "Name", Name);
            string Description = row["DESCRIPTION"].ToString();
            Populate_row(ws, 8, "Description", Description);
            string Implementation_Standards = row["AMPE_GUIDANCE"].ToString();
            Populate_row(ws, 9, "AMPE guidance", Implementation_Standards);
            string Guidance = row["NIST_GUIDANCE"].ToString();
            Populate_row(ws, 10, "NIST guidance", Guidance);
            string Related_Control_Requirements = row["RELATED_CONTROL_REQUIREMENTS"].ToString();
            Populate_row(ws, 11, "Related Control Requirements", Related_Control_Requirements);
            string Prior_ID = row["PRIOR_ID"].ToString();
            Populate_row(ws, 12, "Prior ID", Prior_ID);
#pragma warning restore CS8604 // Possible null reference argument.
#pragma warning restore CS8600 // Converting null literal or possible null value to non-nullable type.


            // Add some style
            ws.Column(1).Width = 40.71;
            ws.Column(2).Width = 139.29;

            // Make the header
            ws.Cells["A1:B1"].Merge = true;
            ws.Cells["A1:B1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Cells["A1:B1"].Style.Font.Bold = true;
            ws.Cells["A1:B1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["A1:B1"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#909090"));
            ws.Cells["A1:B1"].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
            ws.Cells["A1:B1"].Style.Border.Right.Style = ExcelBorderStyle.Medium;
            ws.Cells["A1:B1"].Value = string.Format("{0} Control Information:", audit_name);
        }

        private static void Populate_row(ExcelWorksheet ws, int row, string label, string value)
        {
            Require.False(null == ws, "broken sheet sent");
#pragma warning disable CS8602 // Dereference of a possibly null reference.
            ws.Cells[row, 1].Value = label;
            ws.Cells[row, 1].Style.Font.Bold = true;
            if(value.Contains("{\"richtext\": ["))
            {
                // deal with pretty text
                UglyBlob v = ugly.unpack_pretty(value);
                foreach(UglySegment r in v.segments)
                {
                    ExcelRichText h = ws.Cells[row, 2].RichText.Add(ugly.Base64Decode(r.encoded_text));
                    h.Bold = r.bold;
                    h.Italic = r.italic;
                    h.Color = Color.FromArgb(int.Parse(r.color, System.Globalization.NumberStyles.AllowHexSpecifier));
                    h.PreserveSpace = r.preserve_space;
                    h.Strike = r.strike;
                    h.UnderLine = r.underline;
                }
            }
            else ws.Cells[row, 2].Value = value;
            ws.Cells[row, 2].Style.WrapText = true;
            ws.Cells[row, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells[row, 1].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#C0C0C0"));
#pragma warning restore CS8602 // Dereference of a possibly null reference.
        }

        // Enable us to incrementally save the workbook
        public void MultiSave()
        {
            var holdingstream = new MemoryStream();
            Require.False(null == Report, "impossible MultiSave Error");
#pragma warning disable CS8602 // Dereference of a possibly null reference.
            Report.Stream.CopyTo(holdingstream);
#pragma warning restore CS8602 // Dereference of a possibly null reference.
            Report.SaveAs(output);
            holdingstream.SetLength(0);
            Report.Stream.Position = 0;
            Report.Stream.CopyTo(holdingstream);
            Report.Load(holdingstream);
        }
    }
}
