using System.Data;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;
using OfficeOpenXml.DataValidation;

namespace CMS_Control_Page_Generator
{
    class CMS_Control_Page_Generator
    {
        public static void Add_Control(ExcelWorksheet ws, DataRow control, bool audit)
        {
            ws.Column(1).Width = 40.71;
            ws.Column(2).Width = 139.29;

            Add_Control_Title(ws, control);
            Control_Information(ws, control);
            Instructions(ws);
            Preaudit_Review(ws, control, audit);
            Audit_Results(ws, control, audit);
            Insert_Test_Templates(ws);
        }

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
            ws.Cells["A4:A10"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            ws.Cells["A4:A10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            ws.Cells["A4:A10"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["A4:A10"].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#C0C0C0"));
            ws.Cells["A4:A10"].Style.Font.Bold = true;
            ws.Cells["B4:B10"].Style.WrapText = true;

            // Block values
            ws.Cells[4, 1].Value = "Control Family";
            Set_all_borders(ws, "A4");
            ws.Cells[4, 2].Value = control["FAMILY_ID"] + "-" + control["FAMILY_NAME"];
            Set_all_borders(ws, "B4");
            ws.Cells[5, 1].Value = "Control ID (Type)";
            Set_all_borders(ws, "A5");
            ws.Cells[5, 2].Value = control["CONTROL"] + " (" + control["TYPE"] + ")";
            Set_all_borders(ws, "B5");
            ws.Cells[6, 1].Value = "Control Name";
            Set_all_borders(ws, "A6");
            ws.Cells[6, 2].Value = control["NAME"];
            Set_all_borders(ws, "B6");
            ws.Cells[7, 1].Value = "Control Description";
            Set_all_borders(ws, "A7");
            ws.Cells[7, 2].Value = control["DESCRIPTION"];
            ws.Cells[7, 2].Style.WrapText = true;
            Set_all_borders(ws, "B7");
            ws.Cells[8, 1].Value = "Implementation Standards";
            Set_all_borders(ws, "A8");
            ws.Cells[8, 2].Value = control["IMPLEMENTATION_STANDARDS"];
            Set_all_borders(ws, "B8");
            ws.Cells[9, 1].Value = "Guidance";
            Set_all_borders(ws, "A9");
            ws.Cells[9, 2].Value = control["GUIDANCE"];
            Set_all_borders(ws, "B9");
            ws.Cells[10, 1].Value = "Assessment Methods & Objects";
            Set_all_borders(ws, "A10");
            ws.Cells[10, 2].Value = control["ASSESSMENT_METHOD"];
            Set_all_borders(ws, "B10");
        }

        private static void Instructions(ExcelWorksheet ws)
        {
            // Header
            ws.Cells["A12:B12"].Merge = true;
            ws.Cells["A12:B12"].Value = "Instructions for Pre-Audit Review Activities";
            ws.Cells["A12:B12"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Cells["A12:B12"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["A12:B12"].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#808080"));
            ws.Cells["A12:B12"].Style.Font.Bold = true;

            // Block style
            ws.Cells["A13:A19"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            ws.Cells["A13:A19"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            ws.Cells["A13:A22"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["A13:A22"].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#C0C0C0"));
            ws.Cells["A20"].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#FFFFFF"));
            ws.Cells["A13:A22"].Style.Font.Bold = true;
            ws.Cells["A13:B22"].Style.WrapText = true;
            ws.Cells["B14"].Style.Font.Color.SetColor(Color.Red);
            ws.Cells["B16:B22"].Style.Font.Color.SetColor(Color.Red);


            // Block values
            ws.Cells[13, 1].Value = "General:";
            Set_all_borders(ws, "A13");


            // Format General
            ws.Cells[13, 2].IsRichText = true;
            ExcelRichText ert_general = ws.Cells[13, 2].RichText.Add("Complete the section below titled Pre-Audit Review.  Following the instructions to complete the sections highlighted in yellow or gold.\n");
            ert_general.Bold = false;
            ert_general.Color = System.Drawing.Color.Red;
            ert_general.Italic = true;

            ert_general = ws.Cells[13, 2].RichText.Add("Note: Where available entries from a prior audit have been provided.");
            ert_general.Bold = true;
            ert_general.Color = System.Drawing.Color.Red;
            ert_general.Italic = false;
            Set_all_borders(ws, "B13");

            ws.Cells[14, 1].Value = "Related to App or Agency Service";
            Set_all_borders(ws, "A14");
            ws.Cells[14, 2].Value = "If blank or status has changed select from the drop down to indicate whether the control is applicable to your group.  Note: carried to the summary page.";
            Set_all_borders(ws, "B14");
            ws.Cells[15, 1].Value = " Describe how the control is implemented or why it is not applicable.";
            Set_all_borders(ws, "A15");

            // Format Describe
            ws.Cells[15, 2].IsRichText = true;
            ExcelRichText ert_Describe = ws.Cells[15, 2].RichText.Add("If applicable");
            ert_Describe.Bold = true;
            ert_Describe.Color = System.Drawing.Color.Red;
            ert_Describe.Italic = false;

            ert_Describe = ws.Cells[15, 2].RichText.Add(" - Enter or update the description of how this control is implemented including details of any implementation standards.\n");
            ert_Describe.Bold = false;
            ert_Describe.Color = System.Drawing.Color.Red;
            ert_Describe.Italic = true;

            ert_Describe = ws.Cells[15, 2].RichText.Add("If not applicable");
            ert_Describe.Bold = true;
            ert_Describe.Color = System.Drawing.Color.Red;
            ert_Describe.Italic = false;

            ert_Describe = ws.Cells[15, 2].RichText.Add(" - enter or update the explanation for why it does not apply and enter who is responsible if known.");
            ert_Describe.Bold = false;
            ert_Describe.Color = System.Drawing.Color.Red;
            ert_Describe.Italic = true;
            Set_all_borders(ws, "B15");

            ws.Cells[16, 1].Value = "If Applicable complete the following:";
            Set_all_borders(ws, "A16");
            ws.Cells[16, 2].Value = "Complete the following yellow highlighted sections if the control is applicable to your group.";
            ws.Cells[16, 2].Style.Font.Bold = true;
            ws.Cells["B16"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["B16"].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#F0F0F0"));


            ws.Cells[17, 1].Value = " Describe the Evidence that is available";
            Set_all_borders(ws, "A17");
            ws.Cells[17, 2].Value = "Enter a description of the evidence that will be provided to prove that the control description and implementation standards are being met.";
            Set_all_borders(ws, "B17");
            ws.Cells[18, 1].Value = " How will the evidence be provided";
            Set_all_borders(ws, "A18");
            ws.Cells[18, 2].Value = "Select for the drop down to select how the evidence will be provided to the auditors";
            Set_all_borders(ws, "B18");
            ws.Cells[19, 1].Value = " Source of evidence";
            Set_all_borders(ws, "A19");
            ws.Cells[19, 2].Value = "List the source of the evidence to be provided if the source is someone besides the person identified.";
            Set_all_borders(ws, "B19");

            ws.Cells[21, 1].Value = "STATUS";
            Set_all_borders(ws, "A21");
            ws.Cells[21, 2].Value = "Select from the drop down to change the status of this control after Date Description is completed and the date final evidence is submitted.   Note: carried to the summary page.";
            Set_all_borders(ws, "B21");
            ws.Cells[22, 1].Value = "Dates Pre-Audit Review Actions  Completed";
            Set_all_borders(ws, "A22");
            ws.Cells[22, 2].Value = "Dates: 1) when Pre-Audit on this form is complete, 2) when evidence is submitted.   Note: carried to the summary page.";
            Set_all_borders(ws, "B22");
        }

        private static void Preaudit_Review(ExcelWorksheet ws, DataRow control, bool audit)
        {
            // Header
            ws.Cells["A24:B24"].Merge = true;
            ws.Cells["A24:B24"].Value = "Pre-Audit Review";
            ws.Cells["A24:B24"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Cells["A24:B24"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["A24:B24"].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#808080"));
            ws.Cells["A24:B24"].Style.Font.Bold = true;

            // Block style
            ws.Cells["A25:A34"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            ws.Cells["A25:A34"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            ws.Cells["A25:B34"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["A25:A30"].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#C0C0C0"));
            ws.Cells["B25:B30"].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#FFFF00"));
            ws.Cells["A32:A34"].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#C0C0C0"));
            ws.Cells["B32:B34"].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#FFC000"));
            ws.Cells["A25:A34"].Style.Font.Bold = true;
            ws.Cells["A25:B34"].Style.WrapText = true;

            // Block values
            ws.Cells[25, 1].Value = "Related to App or Agency Service";
            Set_all_borders(ws, "A25");
            ws.Cells[25, 2].Value = control["RELATED_TO"];
            Set_all_borders(ws, "B25");
            ws.Cells[26, 1].Value = " Description of how the control is implement- ed or why not applicable. (Enter/Update)";
            Set_all_borders(ws, "A26");
            ws.Cells[26, 2].Value = control["IMPLEMENTATION"];
            Set_all_borders(ws, "B26");
            ws.Cells[27, 1].Value = "If Applicable complete the following:";
            Set_all_borders(ws, "A27");
            ws.Cells[27, 2].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#C0C0C0"));
            ws.Cells[28, 1].Value = " Describe the Evidence that is available";
            Set_all_borders(ws, "A28");
            ws.Cells[28, 2].Value = control["EVIDENCE_DESCRIPTION"];
            Set_all_borders(ws, "B28");
            ws.Cells[29, 1].Value = " How will the evidence be provided";
            Set_all_borders(ws, "A29");
            ws.Cells[29, 2].Value = control["EVIDENCE_DELIVERY"];
            Set_all_borders(ws, "B29");
            ws.Cells[30, 1].Value = " Source of evidence";
            Set_all_borders(ws, "A30");
            ws.Cells[30, 2].Value = control["EVIDENCE_FILES"];
            Set_all_borders(ws, "B30");

            ws.Cells["A31:B31"].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#FFFFFF"));

            var Preaudit_review_Related = ws.DataValidations.AddListValidation("B25");
            Preaudit_review_Related.Formula.Values.Add("Applicable");
            Preaudit_review_Related.Formula.Values.Add("Not Applicable");

            var Preaudit_review_How = ws.DataValidations.AddListValidation("B29");
            Preaudit_review_How.Formula.Values.Add("Submission");
            Preaudit_review_How.Formula.Values.Add("Demo");
            Preaudit_review_How.Formula.Values.Add("Both");
            Preaudit_review_How.Formula.Values.Add("Other");

            var Preaudit_review_Status = ws.DataValidations.AddListValidation("B32");
            Preaudit_review_Status.Formula.Values.Add("In Process");
            Preaudit_review_Status.Formula.Values.Add("Complete");

            ws.Cells[32, 1].Value = "STATUS of Pre-Audit Review";
            Set_all_borders(ws, "A32");
            if (!audit) ws.Cells[32, 2].Value = control["STATUS"];
            Set_all_borders(ws, "B32");

            var validdateTime33 = ws.DataValidations.AddDateTimeValidation("B33");
            validdateTime33.ShowErrorMessage = true;
            validdateTime33.ErrorStyle = ExcelDataValidationWarningStyle.stop;
            validdateTime33.Error = "Invalid date, must be no earlier than 6 months prior or 3 months after today";
            validdateTime33.ShowInputMessage = true;
            validdateTime33.ShowErrorMessage = true;
            validdateTime33.Prompt = "Must Enter a date";
            validdateTime33.Operator = ExcelDataValidationOperator.between;
            validdateTime33.Formula.Value = DateTime.Today.AddMonths(-6);
            validdateTime33.Formula2.Value = DateTime.Today.AddMonths(4);
            validdateTime33.AllowBlank = true;

            ws.Cells[33, 1].Value = "Date Desc of Implementation Complete";
            Set_all_borders(ws, "A33");
            if (!audit) ws.Cells[33, 2].Value = control["PREAUDIT_REVIEW_DATE"];
            Set_all_borders(ws, "B33");
            ws.Cells[33, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            ws.Cells[33, 2].Style.Numberformat.Format = "MM/DD/YYYY";

            var validdateTime34 = ws.DataValidations.AddDateTimeValidation("B34");
            validdateTime34.ShowErrorMessage = true;
            validdateTime34.ErrorStyle = ExcelDataValidationWarningStyle.stop;
            validdateTime34.Error = "Invalid date, must be no earlier than 6 months prior or 3 months after today";
            validdateTime34.ShowInputMessage = true;
            validdateTime34.ShowErrorMessage = true;
            validdateTime34.Prompt = "Must Enter a date";
            validdateTime34.Operator = ExcelDataValidationOperator.between;
            validdateTime34.Formula.Value = DateTime.Today.AddMonths(-6);
            validdateTime34.Formula2.Value = DateTime.Today.AddMonths(4);
            validdateTime34.AllowBlank = true;


            ws.Cells[34, 1].Value = "Date Evidence Submitted";
            Set_all_borders(ws, "A34");
            if (!audit) ws.Cells[34, 2].Value = control["EVIDENCE_SUBMIT_DATE"];
            Set_all_borders(ws, "B34");
            ws.Cells[34, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            ws.Cells[34, 2].Style.Numberformat.Format = "MM/DD/YYYY";
        }

        private static void Audit_Results(ExcelWorksheet ws, DataRow control, bool audit)
        {
            // First Header
            ws.Cells["A37:B37"].Merge = true;
            ws.Cells["A37:B37"].Value = "Audit Results";
            ws.Cells["A37:B37"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Cells["A37:B37"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["A37:B37"].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#808080"));
            ws.Cells["A37:B37"].Style.Font.Bold = true;

            ws.Cells["A38:A39"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["A38:A39"].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#C0C0C0"));
            ws.Cells[38, 1].Value = "Previous Audit - Decription";
            Set_all_borders(ws, "A38");
            ws.Cells[38, 2].Value = control["LAST_AUDIT"];
            ws.Cells[38, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            Set_all_borders(ws, "B38");
            ws.Cells[39, 1].Value = "Previous Audit - Result:";
            Set_all_borders(ws, "A39");
            ws.Cells[39, 2].Value = control["PREAUDIT_STATUS"];
            Set_all_borders(ws, "B39");

            // Second Header
            ws.Cells["A43:B43"].Merge = true;
            ws.Cells["A43:B43"].Value = "Current Audit";
            ws.Cells["A43:B43"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Cells["A43:B43"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["A43:B43"].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#808080"));
            ws.Cells["A43:B43"].Style.Font.Bold = true;

            ws.Cells["A44:B44"].Merge = true;
            ws.Cells["A44:B44"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["A44:B44"].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#A0A0A0"));
            ws.Cells["A45:A46"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["A45:A46"].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#C0C0C0"));
            ws.Cells["A44:B44"].Value = "Final Result of Audit";
            ws.Cells["A44:B44"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Cells["A44:B44"].Style.Font.Bold = true;

            var review_Status = ws.DataValidations.AddListValidation("B45");
            review_Status.Formula.Values.Add("Gap");
            review_Status.Formula.Values.Add("No Gap");
            review_Status.Formula.Values.Add("Follow up");
            review_Status.Formula.Values.Add("Under Review");

            var criticality_status = ws.DataValidations.AddListValidation("B46");
            criticality_status.Formula.Values.Add("High");
            criticality_status.Formula.Values.Add("Moderate");
            criticality_status.Formula.Values.Add("Low");

            ws.Cells[45, 1].Value = "Status";
            if (!audit) ws.Cells[45, 2].Value = control["AUDIT_STATUS"];
            Set_all_borders(ws, "B45");
            ws.Cells[46, 1].Value = "Criticality";
            if (!audit) ws.Cells[46, 2].Value = control["CRITICALITY"];
            Set_all_borders(ws, "B46");
        }

        private static void Insert_Test_Templates(ExcelWorksheet ws)
        {
            ws.Cells["A48:B48"].Merge = true;
            ws.Cells["A48:B48"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["A48:B48"].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#A0A0A0"));
            ws.Cells["A48:B48"].Value = "Tests Performed";
            Set_all_borders(ws, "A48:B48");
            ws.Cells["A48:B48"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Cells["A48:B48"].Style.Font.Bold = true;
            Insert_Test(ws, "tst 1.1", 49);
            Insert_Test(ws, "tst 1.2", 57);
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

            ws.Cells[row + 5, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells[row + 5, 1].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#A0A0A0"));
            ws.Cells[row + 5, 1].Value = "Notes/Evidence";
            Set_all_borders(ws, ws.Cells[row + 5, 2].Address);

            ws.Cells[row + 6, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells[row + 6, 1].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#A0A0A0"));
            ws.Cells[row + 6, 1].Value = "Criticality";
            Set_all_borders(ws, ws.Cells[row + 6, 2].Address);

            var review_Status = ws.DataValidations.AddListValidation(ws.Cells[row + 4, 2].Address);
            review_Status.Formula.Values.Add("Gap");
            review_Status.Formula.Values.Add("No Gap");
            review_Status.Formula.Values.Add("Follow up");
            review_Status.Formula.Values.Add("Under Review");

            var criticality_status = ws.DataValidations.AddListValidation(ws.Cells[row + 6, 2].Address);
            criticality_status.Formula.Values.Add("High");
            criticality_status.Formula.Values.Add("Moderate");
            criticality_status.Formula.Values.Add("Low");
        }

        private static void Set_all_borders(ExcelWorksheet ws, string block)
        {
            ws.Cells[block].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            ws.Cells[block].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            ws.Cells[block].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            ws.Cells[block].Style.Border.Top.Style = ExcelBorderStyle.Thin;
        }
    }
}
