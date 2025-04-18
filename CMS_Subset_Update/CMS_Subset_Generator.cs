using System.Data;
using common;


namespace CMS_Subset_Update
{
    class CMS_Subset_Generator
    {
        public SQLiteDatabase? mydatabase;

        public class Agency_Control
        {
            public string? Control;
            public string? Agency;
        }

        private List<Agency_Control>? update;
        private bool upgrade;

        public void Initialize(string Database, bool b)
        {
            Require.True(File.Exists(Database), "Specified database doesn't exist");
            mydatabase = new SQLiteDatabase(Database);
            update = new List<Agency_Control>();
            upgrade = b;
        }

        public int Initialize(string Database, string Agency, bool b)
        {
            Initialize(Database, b);
 
            try
            {
                string sql = string.Format("SELECT AGENCY_ID FROM AGENCIES WHERE NICK_ID = '{0}'", Agency);
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
                Console.WriteLine("unable to find AGENCY_ID for provided NICK_ID");
                Environment.Exit(1);
            }
            // Should never be hit
            return 0;
        }

        public void Generate(int Year)
        {
            if (null == update) return;
            foreach (Agency_Control i in update)
            {
                Insert_Audit_Control(i, Year);
            }
        }

        public void Get_Agency_Controls(int agency, List<string> controls, int version)
        {
            if (null == mydatabase) return;
            if(null == update) return;
            foreach (string i in controls)
            {
                string sql = string.Format("SELECT CONTROL_ID FROM CONTROLS WHERE CONTROL = '{0}' And AUDIT_TYPE = '{1}';", i, version);
                string controlID = mydatabase.ExecuteScalar(sql);
                Agency_Control temp = new()
                {
                    Agency = agency.ToString(),
                    Control = controlID
                };

                // Add into our update list
                update.Add(temp);
            }
        }

        public void Insert_Audit_Control(Agency_Control control, int year)
        {
            if(null == mydatabase) return;
            string ControlID = control.Control;
            Require.False(Match.CaseInsensitive("", ControlID), string.Format("No such control ID: {0}", control.Control));
            string sql;
            string prior = mydatabase.ExecuteScalar(string.Format("SELECT PRIOR_ID FROM CONTROLS WHERE CONTROL_ID = {0};", ControlID));

            if (upgrade && !(Match.CaseInsensitive(prior, string.Empty) || Match.CaseInsensitive(prior, "0")))
            {
                sql = string.Format("SELECT AUDIT, RELATED_TO, IMPLEMENTATION, EVIDENCE_DESCRIPTION, EVIDENCE_DELIVERY, EVIDENCE_FILES, STATUS, CRITICALITY FROM AUDIT_CONTROLS WHERE ((CONTROL_ID = {0}) OR (CONTROL_ID = {1})) AND AGENCY_ID = ({2}) ORDER BY AUTO_ID DESC LIMIT 1;"
                                   , ControlID
                                   , prior
                                   , control.Agency);
            }
            else
            {
                sql = string.Format("SELECT AUDIT, RELATED_TO, IMPLEMENTATION, EVIDENCE_DESCRIPTION, EVIDENCE_DELIVERY, EVIDENCE_FILES, STATUS, CRITICALITY FROM AUDIT_CONTROLS WHERE CONTROL_ID = {0} AND AGENCY_ID = ({1}) ORDER BY AUTO_ID DESC LIMIT 1;"
                                   , ControlID
                                   , control.Agency);
            }

            DataTable temp = mydatabase.GetDataTable(sql);

            if (0 == temp.Rows.Count)
            { // First year for the control assume needs work
                sql = String.Format("INSERT INTO AUDIT_CONTROLS (AUDIT, CONTROL_ID, AGENCY_ID, RELATED_TO, PREAUDIT_STATUS, CRITICALITY) VALUES ({0}, '{1}', '{2}', 'Applicable', '','' );"
                                   , year.ToString()
                                   , ControlID
                                   , control.Agency);
                mydatabase.ExecuteNonQuery(sql);
                Console.WriteLine(string.Format("Populating {0} :: {1}", control.Agency, control.Control));
            }
#pragma warning disable CS8604 // Possible null reference argument.
            else if (Match.CaseInsensitive(year.ToString(), temp.Rows[0]["AUDIT"].ToString()))
            { // Previously Generated, stop
                Console.WriteLine(string.Format("Previously Generated {0} :: {1}", control.Agency, control.Control));
                return;
            }
#pragma warning restore CS8604 // Possible null reference argument.
            else
            {
#pragma warning disable CS8600 // Converting null literal or possible null value to non-nullable type.
                string Year = temp.Rows[0]["AUDIT"].ToString();
                Require.False(null == Year, "impossible year");
                // From now on, they should be applicable by default
                string Related_To = "Applicable";//temp.Rows[0]["RELATED_TO"].ToString();
                Require.False(null == Related_To, "impossible Related_To");
                string IMPLEMENTATION = temp.Rows[0]["IMPLEMENTATION"].ToString();
                Require.False(null == IMPLEMENTATION, "impossible IMPLEMENTATION");
                string EVIDENCE_DESCRIPTION = temp.Rows[0]["EVIDENCE_DESCRIPTION"].ToString();
                Require.False(null == EVIDENCE_DESCRIPTION, "impossible EVIDENCE_DESCRIPTION");
                string EVIDENCE_DELIVERY = temp.Rows[0]["EVIDENCE_DELIVERY"].ToString();
                Require.False(null == EVIDENCE_DELIVERY, "impossible EVIDENCE_DELIVERY");
                string EVIDENCE_FILES = temp.Rows[0]["EVIDENCE_FILES"].ToString();
                Require.False(null == EVIDENCE_FILES, "impossible EVIDENCE_FILES");
                // Our preaudit status is last audit's status
                string Preaudit_Status = temp.Rows[0]["STATUS"].ToString();
                Require.False(null == Preaudit_Status, "impossible Preaudit_Status");
                string Criticality = temp.Rows[0]["CRITICALITY"].ToString();
                Require.False(null == Criticality, "impossible Criticality");
#pragma warning disable CS8604 // Possible null reference argument.
                if (!Match.CaseInsensitive("Not Applicable", Related_To))
                {
                    if (Match.CaseInsensitive("", Preaudit_Status))
                    {
                        Preaudit_Status = "Gap";
                    }
                }
#pragma warning restore CS8604 // Possible null reference argument.
#pragma warning restore CS8600 // Converting null literal or possible null value to non-nullable type.
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                sql = String.Format("INSERT INTO AUDIT_CONTROLS (AUDIT, CONTROL_ID, AGENCY_ID, RELATED_TO, IMPLEMENTATION, EVIDENCE_DESCRIPTION, EVIDENCE_DELIVERY, EVIDENCE_FILES, PREAUDIT_STATUS, CRITICALITY, LAST_AUDIT) VALUES ({0}, '{1}', '{2}', '{3}', '{4}', '{5}', '{6}', '{7}', '{8}', '{9}', '{10}' );"
                                   , year.ToString()
                                   , control.Control
                                   , control.Agency
                                   , Related_To.Replace("'", "''")
                                   , IMPLEMENTATION.Replace("'", "''")
                                   , EVIDENCE_DESCRIPTION.Replace("'", "''")
                                   , EVIDENCE_DELIVERY.Replace("'", "''")
                                   , EVIDENCE_FILES.Replace("'", "''")
                                   , Preaudit_Status.Replace("'", "''")
                                   , Criticality.Replace("'", "''")
                                   , Year);
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                mydatabase.ExecuteNonQuery(sql);
            }
        }
    }
}
