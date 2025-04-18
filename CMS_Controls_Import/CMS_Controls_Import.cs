using common;
using OfficeOpenXml;

namespace CMS_Controls_Import
{
    public class CMS_Controls_Import
    {
        public SQLiteDatabase? mydatabase;
        private ExcelWorkbook? book;

        public string Initialize(string Database, FileInfo existingFile, string Audit_Name)
        {
            book = new ExcelPackage(existingFile).Workbook;

            // Connect to Database
            mydatabase = new SQLiteDatabase(Database);

            Require.False(null == mydatabase, string.Format("Unable to open Database file: {0}", Database));

            string sql = string.Format("SELECT AUDIT_TYPE FROM AUDIT_TYPES WHERE UPPER(NAME) LIKE UPPER('{0}');", Audit_Name);
#pragma warning disable CS8602 // Dereference of a possibly null reference.
            string prev = mydatabase.ExecuteScalar(sql);
#pragma warning restore CS8602 // Dereference of a possibly null reference.

            if (Match.CaseInsensitive("",prev) || Match.CaseInsensitive(string.Empty,prev))
            {
                // Entirely new entry
                string sql2 = string.Format("INSERT INTO AUDIT_TYPES(NAME) VALUES('{0}');", Audit_Name);
                mydatabase.ExecuteNonQuery(sql2);
                prev = mydatabase.ExecuteScalar(sql);
                Console.WriteLine("New entry added: {0}\nassigned ID: {1}", Audit_Name, prev);
            }
            else
            {
                // Update Old Entry
                string sql2 = string.Format("DELETE FROM CONTROLS WHERE CONTROLS.AUDIT_TYPE = {0};", prev);
                mydatabase.ExecuteNonQuery(sql2);
                Console.WriteLine("Entry {0} Found: {1}\nReset", Audit_Name, prev);
            }

            return prev;

        }

        public void Cleanup()
        {
            Require.False(null == mydatabase, "impossible cleanup state hit");
#pragma warning disable CS8602 // Dereference of a possibly null reference.
            mydatabase.SQLiteDatabase_Close();
#pragma warning restore CS8602 // Dereference of a possibly null reference.
        }

        public void Import(string AuditType, string AuditName)
        {
            Require.False(null == book, "Unable to read excel input");
            if (null == mydatabase) return;
            if (null == book) return;
            if (Match.CaseInsensitive("CMS ARC-AMPE", AuditName)) Import_arc_ampe(AuditType);
            else if (Match.CaseInsensitive("CMS MARS-E 2.2", AuditName)) Import_CMS(AuditType);
            else if (Match.CaseInsensitive("CMS MARS-E 2.0", AuditName)) Import_CMS(AuditType);
            else
            {
                Console.WriteLine(string.Format("Unknown Control Type", AuditName));
                Environment.Exit(1);
            }
        }

        private void Import_CMS(string AuditType)
        {
            foreach (ExcelWorksheet ws in book.Worksheets)
            {
                string Control_ID = Read_field(ws, "A2", "Control ID", "B2", true,false);
                string Family_ID = Read_field(ws, "A3", "Family ID", "B3", true, false);
                string Family_Name = Read_field(ws, "A4", "Family Name", "B4", true, false);
                string Control = Read_field(ws, "A5", "Control", "B5", true, false);
                string Type = Read_field(ws, "A6", "Type", "B6", true, false);
                string Name = Read_field(ws, "A7", "Name", "B7", true, false);
                string Description = Read_field(ws, "A8", "Description", "B8", true, false);
                string Implementation_Standards = Read_field(ws, "A9", "Implementation Standards", "B9", false, false);
                string Guidance = Read_field(ws, "A10", "Guidance", "B10", false, false);
                string Related_Control_Requirements = Read_field(ws, "A11", "Related Control Requirements", "B11", false, false);
                string Assessment_Objective = Read_field(ws, "A12", "Assessment Objective", "B12", false, false);
                string Assessment_Method = Read_field(ws, "A13", "Assessment Method", "B13", false, false);
                string Prior_ID = Read_field(ws, "A14", "Prior ID", "B14", false, false);

                Dictionary<string, string> tmp = new() {{"CONTROL_ID", Control_ID},
                                                        {"FAMILY_ID", Family_ID},
                                                        {"FAMILY_NAME", Family_Name},
                                                        {"CONTROL", Control },
                                                        {"TYPE", Type },
                                                        {"NAME", Name },
                                                        {"DESCRIPTION", Description },
                                                        {"IMPLEMENTATION_STANDARDS", Implementation_Standards },
                                                        {"GUIDANCE", Guidance },
                                                        {"RELATED_CONTROL_REQUIREMENTS", Related_Control_Requirements },
                                                        {"ASSESSMENT_OBJECTIVE", Assessment_Objective },
                                                        {"ASSESSMENT_METHOD", Assessment_Method },
                                                        {"AUDIT_TYPE", AuditType},
                                                        {"PRIOR_ID", Prior_ID } };
                mydatabase.Insert("CONTROLS", tmp);

            }
        }

        private void Import_arc_ampe(string AuditType)
        {
            foreach (ExcelWorksheet ws in book.Worksheets)
            {
                string Control_ID = Read_field(ws, "A2", "Control ID", "B2", true, false);
                string Family_ID = Read_field(ws, "A3", "Family ID", "B3", true, false);
                string Family_Name = Read_field(ws, "A4", "Family Name", "B4", true, false);
                string Control = Read_field(ws, "A5", "Control", "B5", true, false);
                string Type = Read_field(ws, "A6", "Type", "B6", true, false);
                string Name = Read_field(ws, "A7", "Name", "B7", true, false);
                string Description = Read_field(ws, "A8", "Description", "B8", true, true);
                string AMPE_GUIDANCE = Read_field(ws, "A9", "AMPE guidance", "B9", false, true);
                string NIST_GUIDANCE = Read_field(ws, "A10", "NIST guidance", "B10", false, true);
                string Related_Control_Requirements = Read_field(ws, "A11", "Related Control Requirements", "B11", false, false);
                string Prior_ID = Read_field(ws, "A12", "Prior ID", "B12", false, false);

                Dictionary<string, string> tmp = new() {{"CONTROL_ID", Control_ID},
                                                        {"FAMILY_ID", Family_ID},
                                                        {"FAMILY_NAME", Family_Name},
                                                        {"CONTROL", Control },
                                                        {"TYPE", Type },
                                                        {"NAME", Name },
                                                        {"DESCRIPTION", Description },
                                                        {"AMPE_GUIDANCE", AMPE_GUIDANCE },
                                                        {"NIST_GUIDANCE", NIST_GUIDANCE },
                                                        {"RELATED_CONTROL_REQUIREMENTS", Related_Control_Requirements },
                                                        {"AUDIT_TYPE", AuditType},
                                                        {"PRIOR_ID", Prior_ID } };
                mydatabase.Insert("CONTROLS", tmp);
            }
        }

        private static string Read_field(ExcelWorksheet ws, string label_address, string label_content, string value_address, bool mandatory, bool pretty)
        {
            // Sanity check label
            object hold = ws.Cells[label_address].Value;
            Require.False(null == hold, string.Format("On sheet: {0} field label {1} not found in {2}", ws.Name, label_content, label_address));
#pragma warning disable CS8602 // Dereference of a possibly null reference.
#pragma warning disable CS8600 // Converting null literal or possible null value to non-nullable type.
            string label = hold.ToString();
#pragma warning disable CS8604 // Possible null reference argument.
            Require.True(Match.CaseInsensitive(label_content, label), string.Format("On sheet: {0} field label {1} not found in {2}", ws.Name, label_content,label_address ));
#pragma warning restore CS8604 // Possible null reference argument.
#pragma warning restore CS8600 // Converting null literal or possible null value to non-nullable type.
#pragma warning restore CS8602 // Dereference of a possibly null reference.

            // Get our new value
            if (pretty) hold = ugly.make_ugly(ws.Cells[value_address].RichText);
            else hold = ws.Cells[value_address].Value;
            try
            {
                Require.False(null == hold, string.Format("On sheet: {0} field {1} no value found in {2}", ws.Name, label_content, value_address));
#pragma warning disable CS8602 // Dereference of a possibly null reference.
#pragma warning disable CS8600 // Converting null literal or possible null value to non-nullable type.
                string result = hold.ToString();
#pragma warning disable CS8603 // Possible null reference return.
                return result;
#pragma warning restore CS8603 // Possible null reference return.
#pragma warning restore CS8600 // Converting null literal or possible null value to non-nullable type.
#pragma warning restore CS8602 // Dereference of a possibly null reference.
            }
            catch (Exception)
            {
                if (mandatory) throw;
                else return "";
            }
        }
    }
}
