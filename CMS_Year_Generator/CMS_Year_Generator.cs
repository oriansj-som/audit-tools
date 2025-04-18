using System.Data;
using common;

namespace CMS_Year_Generator
{
    class CMS_Year_Generator
    {
        public SQLiteDatabase? mydatabase;

        public class Agency_Control
        {
            public string? Control;
            public string? Agency;
        }

        private List<Agency_Control>? update;
        private bool upgrade;
        private string? Agencies;
        private List<string>? Inherit_Control_From;

        public void Connect(string Database, bool special, string Agency)
        {
            Agencies = Agency;
            upgrade = special;
            Inherit_Control_From = new List<string>();
            mydatabase = new SQLiteDatabase(Database);
            update = new List<Agency_Control>();
        }

        public void Connect(string Database, bool special, string Agency, List<string> presplit)
        {
            Agencies = Agency;
            upgrade = special;
            Inherit_Control_From = presplit;
            mydatabase = new SQLiteDatabase(Database);
            update = new List<Agency_Control>();
        }

        public void Generate(int Year)
        {
            if (null == update) return;

            foreach (Agency_Control i in update)
            {
                Insert_Audit_Control(i, Year);
            }
        }

        private int Determine_cycle(int year)
        {
            if(null == mydatabase) return -1;

            string sql = string.Format(@"
SELECT MAX(VALUE)
FROM REPORT_DETAILS
WHERE NAME = 'CMS::Base_Year'
  AND VALUE <= {0};"
                                      , year);
            string hold = mydatabase.ExecuteScalar(sql);
            Require.True(4 >= hold.Length, "not a valid year found in database");
            try
            {
                return (year - int.Parse(hold))%3;
            }
            catch
            {
                Console.WriteLine(string.Format("unable to find a CMS::Base_Year which is prior to {0}", year));
                Environment.Exit(1);
            }
            return -1;
        }

        public void Populate_report_details(int year)
        {
            if (null == mydatabase) return;

            string Frontpage = string.Empty;
            string Control = string.Empty;

            int cycle = Determine_cycle(year);
            if (0 == cycle)
            {
                Frontpage = year.ToString() + " CMS Internal First of Three Security Audits";
                Control = year.ToString() + " CMS Internal First of Three Security Audits";
            }
            else if (1 == cycle)
            {
                Frontpage = year.ToString() + " CMS Internal Second of Three Security Audits";
                Control = year.ToString() + " CMS Internal Second of Three Security Audits";
            }
            else if (2 == cycle)
            {
                Frontpage = year.ToString() + " CMS Internal Third of Three Security Audits";
                Control = year.ToString() + " CMS Internal Third of Three Security Audits";
            }
            else
            {
                Console.WriteLine("Some weird shit went down in " + year.ToString());
                Environment.Exit(4);
            }
            try
            {
                if (0 == mydatabase.ExecuteIntScalar(string.Format("SELECT count(1) FROM REPORT_DETAILS WHERE NAME = 'Frontpage::Title' AND VERSION = {0}", year)))
                {
                    mydatabase.Insert("REPORT_DETAILS"
                                     , new Dictionary<string, string> {{ "Name"   , "Frontpage::Title" }
                                                                      ,{ "Value"  , Frontpage          }
                                                                      ,{ "Version", year.ToString()    }});
                }
            }
            catch 
            {
                Console.WriteLine(string.Format("failed to add Frontpage::Title for year {0}", year));
            }
            try
            {
                if (0 == mydatabase.ExecuteIntScalar(string.Format("SELECT count(1) FROM REPORT_DETAILS WHERE NAME = 'Control::Title' AND VERSION = {0}", year)))
                {
                    mydatabase.Insert("REPORT_DETAILS"
                                 , new Dictionary<string, string> {{ "Name"   , "Control::Title" }
                                                                  ,{ "Value"  , Control          }
                                                                  ,{ "Version", year.ToString()  }});
                }
            }
            catch
            {
                Console.WriteLine(string.Format("failed to add Control::Title for year {0}", year));
            }
        }

        private int Lookup_AgencyID(string Agency)
        {
            // shutup compiler
            if(null == mydatabase) return -1;

            string sql = string.Format("SELECT AGENCY_ID FROM AGENCIES WHERE UPPER(NICK_ID) LIKE UPPER('{0}');", Agency);
            int AgencyID;

            try
            {
                string hold = mydatabase.ExecuteScalar(sql);
                AgencyID = int.Parse(hold);
            }
            catch
            {
                throw new Exception(string.Format("Unable to find Agency with NICK_ID of {0}", Agency));
            }
            return AgencyID;
        }

        public void Get_Agency_Controls(int year, List<int> controls, bool all)
        {
            // shutup compiler
            if (null == Agencies) return;
            if (null == mydatabase) return;
            if (null == update) return;

            string sql = string.Empty;
            if (0 == controls.Count)
            {
                int cycle = Determine_cycle(year);
                if      (0 == cycle) sql = "SELECT AGENCY_ID, CONTROL_ID FROM AGENCY_CONTROLS WHERE [0] YEAR_1 = 'True';";
                else if (1 == cycle) sql = "SELECT AGENCY_ID, CONTROL_ID FROM AGENCY_CONTROLS WHERE [0] YEAR_2 = 'True';";
                else if (2 == cycle) sql = "SELECT AGENCY_ID, CONTROL_ID FROM AGENCY_CONTROLS WHERE [0] YEAR_3 = 'True';";
                else
                {
                    // We currently only support the CMS audit 3 year cycle
                    Console.WriteLine("Unable to handle: " + cycle.ToString());
                    Environment.Exit(2);
                }
            }
            else
            {
                // user selects the control years they want
                List<string> where = controls.Select(x => string.Format("YEAR_{0} = 'True'", x)).ToList();
                sql = "SELECT AGENCY_ID, CONTROL_ID FROM AGENCY_CONTROLS WHERE [0] " + string.Join(" OR ", where) + ";";
            }

            if(!Match.CaseInsensitive("ALL", Agencies))
            {
                sql = sql.Replace(";", string.Format(" AND AGENCY_ID = {0};", Lookup_AgencyID(Agencies)));
            }

            // limit ourselves to the target subset
            if (all) sql = sql.Replace("[0]", "");
            else sql = sql.Replace("[0]", "RESPONSIBILITY IN ('Implemented', 'Contribute/Shared') AND ");

            DataTable result = mydatabase.GetDataTable(sql);

            foreach (DataRow i in result.Rows)
            {
                Agency_Control temp = new()
                {
                    Agency = i["AGENCY_ID"].ToString(),
                    Control = i["CONTROL_ID"].ToString()
                };

                // Add into our update list
                update.Add(temp);
            }
        }

        public void Insert_Audit_Control(Agency_Control control, int year)
        {
            // shutup compiler
            if (null == mydatabase) return;
            if (null == Inherit_Control_From) return;

            string ControlID = mydatabase.ExecuteScalar(string.Format("SELECT CONTROL_ID FROM CONTROLS WHERE CONTROL_ID = '{0}' ORDER BY AUDIT_TYPE DESC LIMIT 1;", control.Control));
            Require.False(Match.CaseInsensitive("", ControlID), string.Format("No such control ID: {0}", control.Control));
            string sql;
            string prior = mydatabase.ExecuteScalar(string.Format("SELECT PRIOR_ID FROM CONTROLS WHERE CONTROL_ID = {0};", ControlID));

            if (upgrade && !(Match.CaseInsensitive(prior, string.Empty) || Match.CaseInsensitive(prior, "0")))
            {
                sql = string.Format("SELECT AGENCY_ID, AUDIT, RELATED_TO, IMPLEMENTATION, EVIDENCE_DESCRIPTION, EVIDENCE_DELIVERY, EVIDENCE_FILES, STATUS, CRITICALITY FROM AUDIT_CONTROLS WHERE ((CONTROL_ID = {0}) OR (CONTROL_ID = {1})) AND AGENCY_ID = ({2}) ORDER BY AUTO_ID DESC LIMIT 1;"
                                   , ControlID
                                   ,prior
                                   ,control.Agency);
            }
            else
            {
                sql = string.Format("SELECT AGENCY_ID, AUDIT, RELATED_TO, IMPLEMENTATION, EVIDENCE_DESCRIPTION, EVIDENCE_DELIVERY, EVIDENCE_FILES, STATUS, CRITICALITY FROM AUDIT_CONTROLS WHERE CONTROL_ID = {0} AND AGENCY_ID = ({1}) ORDER BY AUTO_ID DESC LIMIT 1;"
                                   , ControlID
                                   ,control.Agency);
            }

            if(0 < Inherit_Control_From.Count)
            {
                foreach (string s in Inherit_Control_From)
                {
                    sql = sql.Replace("AGENCY_ID = (", string.Format("AGENCY_ID = ({0}, ", Lookup_AgencyID(s)));
                }
                sql = sql.Replace("AGENCY_ID = (", string.Format("AGENCY_ID IN ("));
                if (1 < Inherit_Control_From.Count) sql = sql.Replace(" DESC LIMIT 1", "");
            }
             
            DataTable temp = mydatabase.GetDataTable(sql);
            Require.True(null != temp, "no database?");

#pragma warning disable CS8602 // Dereference of a possibly null reference.
#pragma warning disable CS8604 // Possible null reference argument.
            if (0 == temp.Rows.Count)
            { // First year for the control assume needs work
                sql = String.Format("INSERT INTO AUDIT_CONTROLS (AUDIT, CONTROL_ID, AGENCY_ID, RELATED_TO, PREAUDIT_STATUS, CRITICALITY) VALUES ({0}, '{1}', '{2}', 'Applicable', '','' );"
                                   , year.ToString()
                                   , ControlID
                                   , control.Agency);
                mydatabase.ExecuteNonQuery(sql);
                Console.WriteLine(string.Format("Populating {0} :: {1}", control.Agency, control.Control));
            }
#pragma warning disable CS8602 // Dereference of a possibly null reference.
            else if (Match.CaseInsensitive(year.ToString(), temp.Rows[0]["AUDIT"].ToString()))
            { // Previously Generated, stop
                Console.WriteLine(string.Format("Previously Generated {0} :: {1}", control.Agency, control.Control));
                return;
            }
            else if (1 < temp.Rows.Count)
            {
                merge_agencies(temp, control, year);
            }
            else
            {
#pragma warning disable CS8600 // Converting null literal or possible null value to non-nullable type.
                string Previous_Audit_Year = temp.Rows[0]["AUDIT"].ToString();
                Require.False(null == Previous_Audit_Year, "impossible year");
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
#pragma warning restore CS8600 // Converting null literal or possible null value to non-nullable type.
                if (!Match.CaseInsensitive("Not Applicable", Related_To))
                {
                    if (Match.CaseInsensitive("", Preaudit_Status))
                    {
                        Preaudit_Status = "Gap";
                    }
                }
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
                                   , Previous_Audit_Year);
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                mydatabase.ExecuteNonQuery(sql);
            }
#pragma warning restore CS8604 // Possible null reference argument.
        }

        private void merge_agencies(DataTable temp, Agency_Control control, int year)
        {
            string final_implementation = string.Empty;
            List<string> agencies = new List<string>();
            foreach(DataRow dr in temp.Rows)
            {
                string agency = dr["AGENCY_ID"].ToString();
                if (agencies.Contains(agency)) continue;
                agencies.Add(agency);
                string IMPLEMENTATION = dr["IMPLEMENTATION"].ToString();
                final_implementation = final_implementation + IMPLEMENTATION + "\n";
            }

            string sql = String.Format("INSERT INTO AUDIT_CONTROLS (AUDIT, CONTROL_ID, AGENCY_ID, IMPLEMENTATION, RELATED_TO, PREAUDIT_STATUS, CRITICALITY) VALUES ({0}, '{1}', '{2}', '{3}', 'Applicable', '','' );"
                   ,year.ToString()
                   ,control.Control
                   ,control.Agency
                   ,final_implementation.Replace("'", "''"));
            mydatabase.ExecuteNonQuery(sql);
            Console.WriteLine(string.Format("Populating {0} :: {1}", control.Agency, control.Control));
        }
    }
}
