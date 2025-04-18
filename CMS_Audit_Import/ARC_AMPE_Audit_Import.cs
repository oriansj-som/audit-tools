﻿using OfficeOpenXml;
using common;

namespace CMS_Audit_Import
{
    class ARC_AMPE_Audit_Import
    {
        public static void ARC_AMPE_Import(SQLiteDatabase mydatabase, string Agency, ExcelWorkbook? book, int? Year, int? Audit_Version)
        {
            if (null == mydatabase) return;
            if (null == Agency) return;
            if (null == book) return;
            if (null == Year) return;
            if (null == Audit_Version) return;
            ExcelWorksheet currentWorksheet = book.Worksheets.First();
#pragma warning disable CS8602 // Dereference of a possibly null reference.
            foreach (ExcelWorksheet sheet in book.Worksheets)
            {
                // Skip First Sheet
                // And yes, they do sometimes rename the first sheet
                if (currentWorksheet.Equals(sheet)) continue;

                // Get Universal ControlID
                string sql = string.Format("SELECT CONTROL_ID FROM CONTROLS WHERE CONTROL = '{0}' AND AUDIT_TYPE = {1};", sheet.Name, Audit_Version);
                string controlID = mydatabase.ExecuteScalar(sql);

                int auto = GetAuto(Year, Agency, controlID, mydatabase);
                if (-1 == auto)
                {
                    // Deal with edge case of missing entries
                    string LastAudit = mydatabase.ExecuteScalar(String.Format("SELECT AUDIT FROM AUDIT_CONTROLS WHERE CONTROL_ID = '{0}' AND AGENCY_ID = '{1}' ORDER BY AUDIT DESC LIMIT 1;", controlID, Agency));
                    Dictionary<string, string> tmp = new() {{"AUDIT", Year.ToString()},
                                                            {"AGENCY_ID", Agency},
                                                            {"CONTROL_ID", controlID},
                                                            {"LAST_AUDIT", LastAudit}};
                    mydatabase.Insert("AUDIT_CONTROLS", tmp);
                    auto = GetAuto(Year, Agency, controlID, mydatabase);
                }

                Update_audit_control(auto, sheet, mydatabase);
            }
#pragma warning restore CS8602 // Dereference of a possibly null reference.
        }

        private static int GetAuto(int? audit, string AgencyID, string Control, SQLiteDatabase mydatabase)
        {
            if (null == audit) return -1;
            if (null == mydatabase) return -1;
            string sql = string.Format("SELECT AUTO_ID FROM AUDIT_CONTROLS WHERE AUDIT='{0}' AND AGENCY_ID='{1}' AND CONTROL_ID='{2}';", audit, AgencyID, Control);
            string result = mydatabase.ExecuteScalar(sql);
            try
            {
                return Convert.ToInt32(result);
            }
            catch
            {
                return -1;
            }
        }

        private static void Update_audit_control(int AutoID, ExcelWorksheet ws, SQLiteDatabase mydatabase)
        {
            Update_sub_control(AutoID, ws, "B12", "RELATED_TO", mydatabase, false);
            Update_sub_control(AutoID, ws, "B14", "IMPLEMENTATION", mydatabase, true);
            Update_sub_control(AutoID, ws, "B16", "EVIDENCE_DESCRIPTION", mydatabase, false);
            Update_sub_control(AutoID, ws, "B17", "EVIDENCE_DELIVERY", mydatabase, false);
            Update_sub_control(AutoID, ws, "B18", "EVIDENCE_FILES", mydatabase, false);
            Update_sub_control(AutoID, ws, "B20", "AUDIT_STATUS", mydatabase, false);
            Update_sub_control(AutoID, ws, "B21", "PREAUDIT_REVIEW_DATE", mydatabase, false);
            Update_sub_control(AutoID, ws, "B22", "EVIDENCE_SUBMIT_DATE", mydatabase, false);
            Update_sub_control(AutoID, ws, "B26", "PREAUDIT_STATUS", mydatabase, false);
            Update_sub_control(AutoID, ws, "B30", "STATUS", mydatabase, false);
            Update_sub_control(AutoID, ws, "B31", "CRITICALITY", mydatabase, false);
        }

        private static void Update_sub_control(int AutoID, ExcelWorksheet ws, string Address, string field, SQLiteDatabase mydatabase, bool pretty)
        {
            string current = Read_cell(Address, ws, pretty);
            string sql = string.Format("SELECT {0} FROM Audit_Controls WHERE AUTO_ID='{1}';", field, AutoID);
            Require.False(null == mydatabase, "impossible update_sub_control hit");
#pragma warning disable CS8602 // Dereference of a possibly null reference.
            string data = mydatabase.ExecuteScalar(sql);
#pragma warning restore CS8602 // Dereference of a possibly null reference.

            // Update only if NOT the same
            if (!Match.CaseSensitive(current, data))
            {
                sql = string.Format("UPDATE AUDIT_CONTROLS SET {0} = '{1}' WHERE AUTO_ID='{2}';", field, current.Replace("'", "''"), AutoID);
                int updated = mydatabase.ExecuteNonQuery(sql);
                Require.True(1 == updated, "more than 1 record updated on Audit_Controls");
            }
        }

        private static string Read_cell(string address, ExcelWorksheet currentWorksheet, bool pretty)
        {
            if (pretty) return ugly.make_ugly(currentWorksheet.Cells[address].RichText);
            else return currentWorksheet.Cells[address].Text.Replace('"', '\'');
        }
    }
}
