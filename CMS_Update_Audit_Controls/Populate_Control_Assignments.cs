//comment out if you don't want to remove the old agency_controls
#define clean_sweep

using common;
using OfficeOpenXml;

namespace CMS_Update_Audit_Controls
{
    public struct assignments
    {
        public string agency_nick;
        public string control_name;
    }
    public class Populate_Control_Assignments
    {
        private SQLiteDatabase? mydatabase;
        private int AuditType;
        public List<string>? unknowns;

        public List<assignments> Process(string Database, FileInfo existingFile, string Audit_Name)
        {
            unknowns = new List<string>();
            ExcelWorkbook book = new ExcelPackage(existingFile).Workbook;

            // Connect to Database
            mydatabase = new SQLiteDatabase(Database);

            string sql;
            if (Match.CaseInsensitive(string.Empty, Audit_Name))
            {
                // Use latest audit
                sql = "SELECT MAX(AUDIT_TYPE) FROM AUDIT_TYPES;";
            }
            else
            {
                // Use whatever audit they tell us
                sql = string.Format("SELECT AUDIT_TYPE FROM AUDIT_TYPES WHERE UPPER(NAME) LIKE UPPER('{0}');", Audit_Name);
            }

            try
            {
                string hold = mydatabase.ExecuteScalar(sql);
                AuditType = Convert.ToInt32(hold);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.WriteLine("Unable to find matching Audit name");
                Environment.Exit(1);
            }

            mydatabase = new SQLiteDatabase(Database);

#if clean_sweep
            sql = "DELETE FROM AGENCY_CONTROLS;";
            mydatabase.ExecuteNonQuery(sql);
#endif

            try
            {
                // Get the first worksheet
                ExcelWorksheet currentWorksheet = book.Worksheets.First();
                // Read Agency and Year
                return Read_Page(currentWorksheet);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                Console.WriteLine("Unable to read Excel contents");
                Environment.Exit(1);
            }
            // should never hit
            return null;
        }

        public void Cleanup()
        {
            Require.False(null == mydatabase, "impossible cleanup state hit");
#pragma warning disable CS8602 // Dereference of a possibly null reference.
            mydatabase.SQLiteDatabase_Close();
#pragma warning restore CS8602 // Dereference of a possibly null reference.
        }

        private List<assignments> essential;
        private string agency_name;
        private List<assignments> Read_Page(ExcelWorksheet ws)
        {
            essential = new List<assignments>();
            string columnone = ws.Cells["A10"].Text;
            Require.True(Match.CaseInsensitive("Year", columnone), "Incorrect Column A10 header");
            string columntwo = ws.Cells["B10"].Text;
            Require.True(Match.CaseInsensitive("Cntrl_Id", columntwo), "Incorrect Column B10 header");
            string columnthree = ws.Cells["C10"].Text;
            Require.True(Match.CaseInsensitive("Control", columnthree), "Incorrect Column C10 header");

            int column = 4;
            while (Read_Column(ws, column)) column += 1;
            return essential;
        }

        private bool Read_Column(ExcelWorksheet ws, int column)
        {
            if (null == mydatabase) return false;
            if (null == unknowns) return false;
            string Agency = ws.Cells[10, column].Text;
            if (Match.CaseInsensitive("", Agency)) return false;

            string sql = string.Format("SELECT AGENCY_ID FROM AGENCIES WHERE UPPER(NICK_ID) LIKE UPPER('{0}');", Agency);
            int AgencyID;

            try
            {
                string hold = mydatabase.ExecuteScalar(sql);
                AgencyID = int.Parse(hold);
                agency_name = Agency;
            }
            catch
            {
                unknowns.Add(Agency);
                return true;
            }

            int row = 11;
            while (Read_Row(ws, column, row, AgencyID))
            {
                row += 1;
            }

            return true;
        }

        // memoization to speed up controls
#pragma warning disable IDE0044 // Add readonly modifier
        private static Dictionary<string, int> _controlIDs = new();
#pragma warning restore IDE0044 // Add readonly modifier
        private static int Get_Control_ID(string sql, SQLiteDatabase db, string name)
        {
#pragma warning disable IDE0018 // Inline variable declaration
            int ControlID;
#pragma warning restore IDE0018 // Inline variable declaration
            if (_controlIDs.TryGetValue(sql, out ControlID))
            {
                return ControlID;
            }

            try
            {
                string hold = db.ExecuteScalar(sql);
                ControlID = Convert.ToInt32(hold);
            }
            catch
            {
                Console.WriteLine(string.Format("Control named: {0} does not exist for this audit type", name));
                ControlID = -1;
            }

            _controlIDs[sql] = ControlID;
            return ControlID;
        }

        private bool Read_Row(ExcelWorksheet ws, int column, int row, int AgencyID)
        {
            if (null == mydatabase) return false;
            string control_name = ws.Cells[row, 2].Text.Trim();
            if (Match.CaseInsensitive("", control_name)) return false;
            int controlID = Get_Control_ID(string.Format("SELECT CONTROL_ID FROM CONTROLS WHERE UPPER(CONTROL) LIKE UPPER('{0}') AND AUDIT_TYPE = {1};", control_name, AuditType), mydatabase, control_name);
            // Just skip bad controls
            if (-1 == controlID) return true;

            // add tracking
            assignments a = new assignments();
            a.agency_nick = agency_name;
            a.control_name = control_name;

            string schedule = ws.Cells[row, 1].Text;

            Dictionary<string, string> tmp;

            string applicable = ws.Cells[row, column].Text;
            applicable = applicable.Trim();
            string responsibility = string.Empty;

            if (Match.CaseInsensitive("X", applicable)) responsibility = "Implemented";
            else if (Match.CaseInsensitive("C", applicable)) responsibility = "Contribute/Shared";
            else if (Match.CaseInsensitive("I", applicable)) responsibility = "Inherited";
            else if (Match.CaseInsensitive("N/A", applicable)) return true;
            else if (Match.CaseInsensitive("TBD", applicable))
            {
                Console.WriteLine(string.Format("incomplete record populated on row: {0} column: {1}", row, column));
                return true;
            }
            else if (Match.CaseInsensitive("", applicable)) return true;
            else
            {
                // no idea what this shit is
                Console.WriteLine(string.Format("got {0} on column: {1} row: {2}", applicable, column, row));
                return false;
            }

            if (Match.CaseInsensitive("ALL", schedule))
            {
                tmp = new Dictionary<string, string>() {{"AGENCY_ID", AgencyID.ToString()}
                                                       ,{"CONTROL_ID", controlID.ToString()}
                                                       ,{"RESPONSIBILITY", responsibility }
                                                       ,{"YEAR_1", "True"}
                                                       ,{"YEAR_2", "True"}
                                                       ,{"YEAR_3", "True"}};
            }
            else if (Match.CaseInsensitive("Y1", schedule))
            {
                tmp = new Dictionary<string, string>() {{"AGENCY_ID", AgencyID.ToString()}
                                                       ,{"CONTROL_ID", controlID.ToString()}
                                                       ,{"RESPONSIBILITY", responsibility }
                                                       ,{"YEAR_1", "True"}
                                                       ,{"YEAR_2", "False"}
                                                       ,{"YEAR_3", "False"}};
            }
            else if (Match.CaseInsensitive("Y2", schedule))
            {
                tmp = new Dictionary<string, string>() {{"AGENCY_ID", AgencyID.ToString()}
                                                       ,{"CONTROL_ID", controlID.ToString()}
                                                       ,{"RESPONSIBILITY", responsibility }
                                                       ,{"YEAR_1", "False"}
                                                       ,{"YEAR_2", "True"}
                                                       ,{"YEAR_3", "False"}};
            }
            else if (Match.CaseInsensitive("Y3", schedule))
            {
                tmp = new Dictionary<string, string>() {{"AGENCY_ID", AgencyID.ToString()}
                                                       ,{"CONTROL_ID", controlID.ToString()}
                                                       ,{"RESPONSIBILITY", responsibility }
                                                       ,{"YEAR_1", "False"}
                                                       ,{"YEAR_2", "False"}
                                                       ,{"YEAR_3", "True"}};
            }
            else if (Match.CaseInsensitive("N/A", schedule))
            {
                // Just skip it as we are not covered by this
                return true;
            }
            else
            {
                throw new Exception(string.Format("Unknown pattern: {0}", schedule));
            }

            essential.Add(a);

#if !clean_sweep
            // we just can create duplicates, check first
            string sql = string.Format("SELECT count(1) FROM AGENCY_CONTROLS WHERE AGENCY_ID={0} AND CONTROL_ID={1};", AgencyID, controlID);
            int count = mydatabase.ExecuteIntScalar(sql);
            if (0 == count) mydatabase.Insert("AGENCY_CONTROLS", tmp);
            else mydatabase.Update("AGENCY_CONTROLS", tmp, string.Format("AGENCY_ID={0} AND CONTROL_ID={1}", AgencyID, controlID));
#else
            mydatabase.Insert("AGENCY_CONTROLS", tmp);
#endif
            return true;
        }
    }
}