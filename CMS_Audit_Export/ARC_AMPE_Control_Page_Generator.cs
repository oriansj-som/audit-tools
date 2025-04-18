// If you want cell locking just uncomment
#define cell_locks

using System.Data;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;
using OfficeOpenXml.DataValidation;
using common;
using CMS_Audit_Export;

namespace ARC_AMPE_Control_Page_Generator
{
    class ARC_AMPE_Control_Page_Generator
    {
        public static void Add_Control(ExcelWorksheet ws, DataRow control, bool audit, SQLiteDatabase source, infrastructure infrastructure_platform)
        {
#if cell_locks
            ws.Protection.IsProtected = true;
            ws.Protection.SetPassword(control["CONTROL"].ToString());
#endif
            ws.Column(1).Width = 40.71;
            ws.Column(2).Width = 139.29;
            ws.Column(3).Width = 139.29;

            Add_Control_Title(ws, control);
            Control_Information(ws, control);
            Preaudit_Review(ws, control, audit);
            Audit_Results(ws, control, audit);
            Insert_Test_Templates(ws);
            Add_state_standard_response(ws, control, audit, source, infrastructure_platform);
#if cell_locks
            unlock_bottom(ws, 29, 300);
#endif
        }
#if cell_locks
        private static void unlock_bottom(ExcelWorksheet ws, int start, int end)
        {
            int i = start;
            while (i < end)
            {
                ws.Cells[i, 1].Style.Locked = false;
                ws.Cells[i, 2].Style.Locked = false;
                ws.Cells[i, 2].Style.Locked = false;
                i += 1;
            }
        }
#endif
        private static void Add_Control_Title(ExcelWorksheet ws, DataRow control)
        {
            ws.Cells["A1:B1"].Merge = true;
            ws.Cells["A1:B1"].Value = control["Title"].ToString();
            ws.Cells["A1:B1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Cells["A1:B1"].Style.Font.Bold = true;
            ws.Cells["A1:B1"].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
            ws.Cells["A1:B1"].Style.Border.Right.Style = ExcelBorderStyle.Medium;
        }

        private static void Control_Information(ExcelWorksheet ws, DataRow control)
        {
            // Header
            ws.Cells["A3:B3"].Merge = true;
            ws.Cells["A3:B3"].Value = control["version"].ToString();

            ws.Cells["A3:B3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Cells["A3:B3"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["A3:B3"].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#808080"));
            ws.Cells["A3:B3"].Style.Font.Bold = true;

            // Block style
            ws.Cells["A4:A9"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            ws.Cells["A4:A9"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            ws.Cells["A4:A9"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["A4:A9"].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#C0C0C0"));
            ws.Cells["A4:A9"].Style.Font.Bold = true;
            ws.Cells["B4:B9"].Style.WrapText = true;

            // Block values
            ws.Cells[4, 1].Value = "Control Family";
            Set_all_borders(ws, "A4");
            ws.Cells[4, 2].Value = control["FAMILY_ID"] + "-" + control["FAMILY_NAME"];
            Set_all_borders(ws, "B4");
            ws.Cells[5, 1].Value = "Control Number";
            Set_all_borders(ws, "A5");
            ws.Cells[5, 2].Value = control["CONTROL"];
            Set_all_borders(ws, "B5");
            ws.Cells[6, 1].Value = "Control Name";
            Set_all_borders(ws, "A6");
            ws.Cells[6, 2].Value = control["NAME"];
            Set_all_borders(ws, "B6");
            ws.Cells[7, 1].Value = "Control Description";
            Set_all_borders(ws, "A7");

            string raw_description = control["DESCRIPTION"].ToString();
            Require.True(null != raw_description, "impossible raw description");
#pragma warning disable CS8602 // Dereference of a possibly null reference.
            if (raw_description.Contains("{\"richtext\": ["))
            {
                make_pretty(ws, 7, 2, raw_description);
                //ws.Cells[7, 2].Style.WrapText = true;
                ws.Cells[7, 2].Style.Font.Name = "Consolas";
            }
            else
            {
                ws.Cells[7, 2].Value = raw_description;
                //ws.Cells[7, 2].Style.WrapText = true;
                ws.Cells[7, 2].Style.Font.Name = "Consolas";
            }
#pragma warning restore CS8602 // Dereference of a possibly null reference.
            Set_all_borders(ws, "B7");
            ws.Cells[8, 1].Value = "Supplemental Control Requirements & Guidance";
            ws.Cells[8, 1].Style.WrapText = true;
            Set_all_borders(ws, "A8");

            string raw_amp = control["AMPE_GUIDANCE"].ToString();
            Require.True(null != raw_amp, "impossible raw AMP");
            if (raw_amp.Contains("{\"richtext\": ["))
            {
                make_pretty(ws, 8, 2, raw_amp);
                ws.Cells[8, 2].Style.WrapText = true;
                ws.Cells[8, 2].Style.Font.Name = "Consolas";
            }
            else
            {
                ws.Cells[8, 2].Value = raw_amp;
                ws.Cells[8, 2].Style.WrapText = true;
                ws.Cells[8, 2].Style.Font.Name = "Consolas";
            }
            Set_all_borders(ws, "B8");

            ws.Cells[9, 1].Value = "NIST 800.53 rev 5 Discussion/Guidance";
            ws.Cells[9, 1].Style.WrapText = true;
            Set_all_borders(ws, "A9");

            string raw_nist = control["NIST_GUIDANCE"].ToString();
            Require.True(null != raw_nist, "impossible raw NIST");
            if (raw_nist.Contains("{\"richtext\": ["))
            {
                make_pretty(ws, 9, 2, raw_nist);
                ws.Cells[9, 2].Style.WrapText = true;
                ws.Cells[9, 2].Style.Font.Name = "Consolas";
            }
            else
            {
                ws.Cells[9, 2].Value = raw_nist;
                ws.Cells[9, 2].Style.WrapText = true;
                ws.Cells[9, 2].Style.Font.Name = "Consolas";
            }
            Set_all_borders(ws, "B9");
        }

        private static void Preaudit_Review(ExcelWorksheet ws, DataRow control, bool audit)
        {
            // Header
            ws.Cells["A11:B11"].Merge = true;
            ws.Cells["A11:B11"].Value = "Pre-Audit Review";
            ws.Cells["A11:B11"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Cells["A11:B11"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["A11:B11"].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#808080"));
            ws.Cells["A11:B11"].Style.Font.Bold = true;

            // Block style
            ws.Cells["A12:A22"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            ws.Cells["A12:A22"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            ws.Cells["A12:B22"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["A12:A18"].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#C0C0C0"));
            ws.Cells["B12:B18"].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#FFFF00"));
            ws.Cells["A20:A22"].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#C0C0C0"));
            ws.Cells["B20:B22"].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#FFC000"));
            ws.Cells["A12:A22"].Style.Font.Bold = true;
            ws.Cells["A12:B22"].Style.WrapText = true;

            // Block values
            ws.Cells[12, 1].Value = "Is this control Applicable or Not Applicable.";
            Set_all_borders(ws, "A12");
            ws.Cells[12, 2].Value = control["RELATED_TO"];
#if cell_locks
            ws.Cells[12, 2].Style.Locked = false;
#endif
            Set_all_borders(ws, "B12");

            var Preaudit_review_Related = ws.DataValidations.AddListValidation("B12");
            Preaudit_review_Related.Formula.Values.Add("Applicable");
            Preaudit_review_Related.Formula.Values.Add("Not Applicable");


            ws.Cells[13, 1].Value = "Control Status:";
            Set_all_borders(ws, "A13");
            string responsibility = control["RESPONSIBILITY"].ToString();
            if (!audit && (Match.CaseInsensitive("null", responsibility) || responsibility.Length <= 1))
            {
                Console.WriteLine(string.Format("control: {0} lacks responsibility", control["CONTROL"]));
                Environment.Exit(1);
            }
            else if(audit && (Match.CaseInsensitive("null", responsibility) || responsibility.Length <= 1))
            {
                ws.Cells[13, 2].Value = "Implemented";
            }
            else ws.Cells[13, 2].Value = responsibility;
#if cell_locks
            ws.Cells[13, 2].Style.Locked = false;
#endif
            Set_all_borders(ws, "B13");

            var Preaudit_control_status = ws.DataValidations.AddListValidation("B13");
            Preaudit_control_status.Formula.Values.Add("Contribute/Shared");
            Preaudit_control_status.Formula.Values.Add("Implemented");
            Preaudit_control_status.Formula.Values.Add("Inherited");
            Preaudit_control_status.Formula.Values.Add("Not Applicable");


            ws.Cells[14, 1].Value = "Description of how the control is implemented\r\nOR why it's not applicable.";
            Set_all_borders(ws, "A14");
            string implementation = control["IMPLEMENTATION"].ToString();
            if (implementation.Contains("{\"richtext\": ["))
            {
                make_pretty(ws, 14, 2, implementation);
                ws.Cells[14, 2].Style.Font.Name = "Consolas";
            }
            else ws.Cells[14, 2].Value = implementation;
            ws.Cells[14, 2].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
#if cell_locks
            ws.Cells[14, 2].Style.Locked = false;
#endif
            Set_all_borders(ws, "B14");

            ws.Cells[15, 1].Value = "If Applicable complete the following:";
            Set_all_borders(ws, "A15");
            ws.Cells[15, 2].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#C0C0C0"));
            Set_all_borders(ws, "B15");

            ws.Cells[16, 1].Value = "Describe the Evidence that is available";
            Set_all_borders(ws, "A16");
            ws.Cells[16, 2].Value = control["EVIDENCE_DESCRIPTION"];
#if cell_locks
            ws.Cells[16, 2].Style.Locked = false;
#endif
            Set_all_borders(ws, "B16");

            ws.Cells[17, 1].Value = "How will the evidence be provided";
            Set_all_borders(ws, "A17");
            ws.Cells[17, 2].Value = control["EVIDENCE_DELIVERY"];
#if cell_locks
            ws.Cells[17, 2].Style.Locked = false;
#endif
            Set_all_borders(ws, "B17");

            var Preaudit_review_How = ws.DataValidations.AddListValidation("B17");
            Preaudit_review_How.Formula.Values.Add("Submission");
            Preaudit_review_How.Formula.Values.Add("Demo");
            Preaudit_review_How.Formula.Values.Add("Both");
            Preaudit_review_How.Formula.Values.Add("Other");


            ws.Cells[18, 1].Value = "Filenames of evidence";
            Set_all_borders(ws, "A18");
            ws.Cells[18, 2].Value = control["EVIDENCE_FILES"];
#if cell_locks
            ws.Cells[18, 2].Style.Locked = false;
#endif
            Set_all_borders(ws, "B18");

            ws.Cells["A19:B19"].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#FFFFFF"));

            ws.Cells[20, 1].Value = "STATUS of Pre-Audit Review";
            Set_all_borders(ws, "A20");
            if (!audit) ws.Cells[20, 2].Value = control["STATUS"];
#if cell_locks
            ws.Cells[20,2].Style.Locked = false;
#endif
            Set_all_borders(ws, "B20");

            var Preaudit_review_Status = ws.DataValidations.AddListValidation("B20");
            Preaudit_review_Status.Formula.Values.Add("In Process");
            Preaudit_review_Status.Formula.Values.Add("Complete");

            ws.Cells[21, 1].Value = "Date Desc of Implementation Complete";
            Set_all_borders(ws, "A21");
            if (!audit) ws.Cells[21, 2].Value = control["PREAUDIT_REVIEW_DATE"];
            Set_all_borders(ws, "B21");
            ws.Cells[21, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            ws.Cells[21, 2].Style.Numberformat.Format = "MM/DD/YYYY";
#if cell_locks
            ws.Cells[21, 2].Style.Locked = false;
#endif
            var validdateTime21 = ws.DataValidations.AddDateTimeValidation("B21");
            validdateTime21.ShowErrorMessage = true;
            validdateTime21.ErrorStyle = ExcelDataValidationWarningStyle.stop;
            validdateTime21.Error = "Invalid date, must be no earlier than 6 months prior or 3 months after today";
            validdateTime21.ShowInputMessage = true;
            validdateTime21.ShowErrorMessage = true;
            validdateTime21.Prompt = "Must Enter a date";
            validdateTime21.Operator = ExcelDataValidationOperator.between;
            validdateTime21.Formula.Value = DateTime.Today.AddMonths(-6);
            validdateTime21.Formula2.Value = DateTime.Today.AddMonths(4);
            validdateTime21.AllowBlank = true;

            ws.Cells[22, 1].Value = "Date Evidence Submitted/Uploaded";
            Set_all_borders(ws, "A22");
            if (!audit) ws.Cells[22, 2].Value = control["EVIDENCE_SUBMIT_DATE"];
            Set_all_borders(ws, "B22");
            ws.Cells[22, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            ws.Cells[22, 2].Style.Numberformat.Format = "MM/DD/YYYY";
#if cell_locks
            ws.Cells[22, 2].Style.Locked = false;
#endif
            var validdateTime22 = ws.DataValidations.AddDateTimeValidation("B22");
            validdateTime22.ShowErrorMessage = true;
            validdateTime22.ErrorStyle = ExcelDataValidationWarningStyle.stop;
            validdateTime22.Error = "Invalid date, must be no earlier than 6 months prior or 3 months after today";
            validdateTime22.ShowInputMessage = true;
            validdateTime22.ShowErrorMessage = true;
            validdateTime22.Prompt = "Must Enter a date";
            validdateTime22.Operator = ExcelDataValidationOperator.between;
            validdateTime22.Formula.Value = DateTime.Today.AddMonths(-6);
            validdateTime22.Formula2.Value = DateTime.Today.AddMonths(4);
            validdateTime22.AllowBlank = true;
        }

        private static void Audit_Results(ExcelWorksheet ws, DataRow control, bool audit)
        {
            // First Header
            ws.Cells["A24:B24"].Merge = true;
            ws.Cells["A24:B24"].Value = "Previous Audit Results";
            ws.Cells["A24:B24"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Cells["A24:B24"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["A24:B24"].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#808080"));
            ws.Cells["A24:B24"].Style.Font.Bold = true;

            ws.Cells["A25:A26"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["A25:A26"].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#C0C0C0"));
            ws.Cells[25, 1].Value = "Previous Audit - Year:";
            Set_all_borders(ws, "A25");
            ws.Cells[25, 2].Value = control["LAST_AUDIT"];
            ws.Cells[25, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            Set_all_borders(ws, "B25");

            ws.Cells[26, 1].Value = "Previous Audit - Result:";
            Set_all_borders(ws, "A26");
            ws.Cells[26, 2].Value = control["PREAUDIT_STATUS"];
            Set_all_borders(ws, "B26");

            // Second Header
            ws.Cells["A29:B29"].Merge = true;
            ws.Cells["A29:B29"].Value = "Current Audit Results";
            ws.Cells["A29:B29"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Cells["A29:B29"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["A29:B29"].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#808080"));
            ws.Cells["A29:B29"].Style.Font.Bold = true;

            ws.Cells[30, 1].Value = "Current Audit - Status:";
            Set_all_borders(ws, "A30");
            if (!audit) ws.Cells[30, 2].Value = control["STATUS"];
#if cell_locks
            ws.Cells[30, 2].Style.Locked = false;
#endif
            Set_all_borders(ws, "B30");

            var review_Status = ws.DataValidations.AddListValidation("B30");
            review_Status.Formula.Values.Add("Gap");
            review_Status.Formula.Values.Add("No Gap");
            review_Status.Formula.Values.Add("Follow up");
            review_Status.Formula.Values.Add("Under Review");


            ws.Cells[31, 1].Value = "Current Audit - Criticality:";
            Set_all_borders(ws, "A31");
            if (!audit) ws.Cells[46, 2].Value = control["CRITICALITY"];
#if cell_locks
            ws.Cells[31, 2].Style.Locked = false;
#endif
            Set_all_borders(ws, "B31");
            var criticality_status = ws.DataValidations.AddListValidation("B31");
            criticality_status.Formula.Values.Add("High");
            criticality_status.Formula.Values.Add("Moderate");
            criticality_status.Formula.Values.Add("Low");
        }

        private static void Insert_Test_Templates(ExcelWorksheet ws)
        {
            ws.Cells["A33:B33"].Merge = true;
            ws.Cells["A33:B33"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["A33:B33"].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#A0A0A0"));
            ws.Cells["A33:B33"].Value = "Tests Performed";
            Set_all_borders(ws, "A33:B33");
            ws.Cells["A33:B33"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Cells["A33:B33"].Style.Font.Bold = true;
            Insert_Test(ws, "tst 1.1", 34);
            Insert_Test(ws, "tst 1.2", 42);
        }

        private static void Insert_Test(ExcelWorksheet ws, string number, int row)
        {
            ws.Cells[row, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells[row, 1].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#A0A0A0"));
            ws.Cells[row, 1].Style.Font.Bold = true;
            ws.Cells[row, 1].Value = "Test ID";
            Set_all_borders(ws, ws.Cells[row, 1].Address);

            ws.Cells[row, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells[row, 2].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#A0A0A0"));
            ws.Cells[row, 2].Value = number;

            ws.Cells[row + 1, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells[row + 1, 1].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#A0A0A0"));
            ws.Cells[row + 1, 1].Value = "Test Procedures (Assessment Procedures)";
            Set_all_borders(ws, ws.Cells[row + 1, 2].Address);

            ws.Cells[row + 2, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells[row + 2, 1].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#A0A0A0"));
            ws.Cells[row + 2, 1].Value = "Expected Results";
            Set_all_borders(ws, ws.Cells[row + 2, 2].Address);

            ws.Cells[row + 3, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells[row + 3, 1].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#A0A0A0"));
            ws.Cells[row + 3, 1].Value = "Actual Results";
            Set_all_borders(ws, ws.Cells[row + 3, 2].Address);

            ws.Cells[row + 4, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells[row + 4, 1].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#A0A0A0"));
            ws.Cells[row + 4, 1].Value = "Status";
            Set_all_borders(ws, ws.Cells[row + 4, 2].Address);

            var review_Status = ws.DataValidations.AddListValidation(ws.Cells[row + 4, 2].Address);
            review_Status.Formula.Values.Add("Gap");
            review_Status.Formula.Values.Add("No Gap");
            review_Status.Formula.Values.Add("Follow up");
            review_Status.Formula.Values.Add("Under Review");

            ws.Cells[row + 5, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells[row + 5, 1].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#A0A0A0"));
            ws.Cells[row + 5, 1].Value = "Notes/Evidence";
            Set_all_borders(ws, ws.Cells[row + 5, 2].Address);

            ws.Cells[row + 6, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells[row + 6, 1].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#A0A0A0"));
            ws.Cells[row + 6, 1].Value = "Criticality";
            Set_all_borders(ws, ws.Cells[row + 6, 2].Address);

            var criticality_status = ws.DataValidations.AddListValidation(ws.Cells[row + 6, 2].Address);
            criticality_status.Formula.Values.Add("High");
            criticality_status.Formula.Values.Add("Moderate");
            criticality_status.Formula.Values.Add("Low");
        }
        private static void Add_state_standard_response(ExcelWorksheet ws, DataRow control, bool audit, SQLiteDatabase source, infrastructure infrastructure_platform)
        {
            // infrastructure providers don't get this
            if (Match.CaseInsensitive("SELF", infrastructure_platform.name)) return;

            string control_id = control["CONTROL"].ToString();
            string sql = string.Format(@"
SELECT IMPLEMENTATION
FROM AUDIT_CONTROLS
INNER JOIN CONTROLS
ON AUDIT_CONTROLS.CONTROL_ID = CONTROLS.CONTROL_ID
WHERE AUDIT_CONTROLS.AGENCY_ID='{0}'
AND CONTROL='{1}'
ORDER BY AUDIT_CONTROLS.AUTO_ID DESC LIMIT 1;"
                                        ,infrastructure_platform.ID
                                        ,control_id);
            string implementation = source.ExecuteScalar(sql);
            if (string.IsNullOrEmpty(implementation)) implementation = "There is no standard response at this time";

            ws.Cells[13, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells[13, 3].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#A0A0A0"));
            ws.Cells[13, 3].Value = string.Format("CMS SSPP {0} Standard Response", infrastructure_platform.name);
            Set_all_borders(ws, "C13");

            if (implementation.Contains("{\"richtext\": ["))
            {
                make_pretty(ws, 14, 3, implementation);
                ws.Cells[14, 3].Style.Font.Name = "Consolas";
            }
            else ws.Cells[14, 3].Value = implementation;
            ws.Cells[14, 3].Style.WrapText = true;
            ws.Cells[14, 3].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
            Set_all_borders(ws, "C14");
        }

        private static void Set_all_borders(ExcelWorksheet ws, string block)
        {
            ws.Cells[block].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            ws.Cells[block].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            ws.Cells[block].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            ws.Cells[block].Style.Border.Top.Style = ExcelBorderStyle.Thin;
        }
        private static void make_pretty(ExcelWorksheet ws, int row, int column, string value)
        {
            // deal with pretty text
            UglyBlob v = ugly.unpack_pretty(value);
            foreach (UglySegment r in v.segments)
            {
                ExcelRichText h = ws.Cells[row, column].RichText.Add(ugly.Base64Decode(r.encoded_text));
                h.Bold = r.bold;
                h.Italic = r.italic;
                h.Color = Color.FromArgb(int.Parse(r.color, System.Globalization.NumberStyles.AllowHexSpecifier));
                h.PreserveSpace = r.preserve_space;
                h.Strike = r.strike;
                h.UnderLine = r.underline;
            }
        }
    }
}
