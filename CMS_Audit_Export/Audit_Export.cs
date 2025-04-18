using System.Data;
using OfficeOpenXml;
using common;

namespace CMS_Audit_Export
{
    public struct infrastructure
    {
        public string name;
        public string ID;
    }
    class Audit_Export
    {
        public SQLiteDatabase? mydatabase;
        public ExcelPackage? Report;
        public FileInfo? output;
        public bool fresh;

        public void Initialize(string Database, bool audit)
        {
            fresh = audit;
            Require.True(File.Exists(Database), "Specified database doesn't exist");

            // Create our new temp excel file
            Report = new ExcelPackage();

            // Connect to Database
            mydatabase = new SQLiteDatabase(Database);
        }
        public int Initialize(string Database, string Agency, bool audit)
        {
            Initialize(Database, audit);

            try
            {
                string sql = string.Format(@"
SELECT AGENCY_ID
FROM AGENCIES
WHERE NICK_ID = '{0}'"
                                          ,Agency);
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
            return -1;
        }

        public void Generate(FileInfo O, int AgencyID, int Year, string audit_name)
        {
            output = O;
            if (Match.CaseInsensitive("CMS ARC-AMPE", audit_name)) Generate_ARC_AMPE(AgencyID, Year);
            else if (Match.CaseInsensitive("CMS MARS-E 2.2", audit_name)) Generate_CMS(AgencyID, Year);
            else if (Match.CaseInsensitive("CMS MARS-E 2.0", audit_name)) Generate_CMS(AgencyID, Year);
            else
            {
                Console.WriteLine(string.Format("Unknown Control Type", audit_name));
                Environment.Exit(1);
            }
            MultiSave();
        }

        private void Generate_ARC_AMPE(int AgencyID, int Year)
        {
            int worksheet = 2;
            int row = 21;
            string sql = string.Format(@"
SELECT CONTROLS.CONTROL
      ,FAMILY_ID
      ,FAMILY_NAME
      ,TYPE
      ,NAME
      ,DESCRIPTION
      ,AMPE_GUIDANCE
      ,NIST_GUIDANCE
      ,RELATED_TO
	  ,RESPONSIBILITY
      ,IMPLEMENTATION
      ,EVIDENCE_DESCRIPTION
      ,EVIDENCE_DELIVERY
      ,EVIDENCE_FILES
      ,AUDIT_STATUS
      ,PREAUDIT_STATUS
      ,PREAUDIT_REVIEW_DATE
      ,EVIDENCE_SUBMIT_DATE
      ,STATUS
      ,CRITICALITY
      ,LAST_AUDIT
      ,(SELECT VALUE FROM REPORT_DETAILS WHERE NAME = 'Control::Information' AND VERSION = AUDIT_TYPE) version
      ,(SELECT VALUE FROM REPORT_DETAILS WHERE NAME = 'Control::Title' AND VERSION = {0}) Title
FROM CONTROLS
INNER JOIN AUDIT_CONTROLS ON CONTROLS.CONTROL_ID = AUDIT_CONTROLS.CONTROL_ID
LEFT JOIN AGENCY_CONTROLS
ON AUDIT_CONTROLS.AGENCY_ID = AGENCY_CONTROLS.AGENCY_ID AND AUDIT_CONTROLS.CONTROL_ID = AGENCY_CONTROLS.CONTROL_ID
WHERE AUDIT = '{0}'
AND AUDIT_CONTROLS.AGENCY_ID = '{1}'
ORDER BY Controls.Control;"
                                      , Year
                                      , AgencyID);
            Require.False(null == mydatabase, "impossible error in Generate");
#pragma warning disable CS8602 // Dereference of a possibly null reference.
            DataTable temp = mydatabase.GetDataTable(sql);
            // if you see an error here, make sure you have ALL the REPORT_DETAILS set
            Require.False(null == Report, "impossible error in Generate 2");
            Report.Workbook.Worksheets.Add("Summary");
#pragma warning restore CS8602 // Dereference of a possibly null reference.

            infrastructure infrastructure_platform = new infrastructure();
            infrastructure_platform.name = "SELF";
            sql = string.Format(@"
SELECT AGENCY_ID, NICK_ID
FROM AGENCIES
WHERE AGENCY_ID IN (SELECT PLATFORM_RESPONSE_PROVIDER FROM AGENCIES WHERE AGENCY_ID = {0})
ORDER BY AGENCY_ID ASC LIMIT 1;"
                                 , AgencyID);

            DataTable infra = mydatabase.GetDataTable(sql);
            if (null == infra || 0 == infra.Rows.Count || null == infra.Rows[0] || null == infra.Rows[0]["AGENCY_ID"] || Match.CaseInsensitive(infra.Rows[0]["AGENCY_ID"].ToString(), AgencyID.ToString()))
            {
                infrastructure_platform.name = "SELF";
            }
            else
            {
                infrastructure_platform.name = infra.Rows[0]["NICK_ID"].ToString();
                infrastructure_platform.ID = infra.Rows[0]["AGENCY_ID"].ToString();
            }

            foreach (DataRow i in temp.Rows)
            {
                object? hold = i["CONTROL"];
                Require.False(null == hold, "impossible error in Generate 3");
#pragma warning disable CS8602 // Dereference of a possibly null reference.
#pragma warning disable CS8600 // Converting null literal or possible null value to non-nullable type.
                string name = hold.ToString();
#pragma warning restore CS8600 // Converting null literal or possible null value to non-nullable type.
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                Report.Workbook.Worksheets.Add(name);
                ARC_AMPE_Control_Page_Generator.ARC_AMPE_Control_Page_Generator.Add_Control(Report.Workbook.Worksheets[name], i, true, mydatabase, infrastructure_platform);
                worksheet += 1;
                ARC_AMPE_Front_Page_Generator.ARC_AMPE_Front_Page_Generator.Add_Frontpage_row(Report.Workbook.Worksheets["Summary"], i, row);
                row += 1;
            }

            // Get the audit Type
            sql = string.Format(@"
SELECT MAX(AUDIT_TYPE)
FROM CONTROLS
INNER JOIN AUDIT_CONTROLS ON CONTROLS.CONTROL_ID = AUDIT_CONTROLS.CONTROL_ID
WHERE AUDIT = '{0}' AND AGENCY_ID = '{1}'
ORDER BY CONTROLS.CONTROL;"
                                , Year
                                , AgencyID);
            string version = mydatabase.ExecuteScalar(sql);

            // Get page Title
            sql = string.Format(@"
SELECT VALUE
FROM REPORT_DETAILS
WHERE NAME = 'Frontpage::Title'
  AND VERSION = {0}"
                               , Year);
            string Title = mydatabase.ExecuteScalar(sql);

            // Get Agency info
            sql = string.Format(@"
SELECT AGENCY_ID
      ,TEAM_ID
      ,SYSTEM_ID
      ,PLATFORM
FROM AGENCIES
WHERE AGENCY_ID = '{0}';"
                               , AgencyID);
            temp = mydatabase.GetDataTable(sql);

            ARC_AMPE_Front_Page_Generator.ARC_AMPE_Front_Page_Generator.Add_Frontpage_Header(Report.Workbook.Worksheets["Summary"], row - 1, temp, Title, version);
        }

        private void Generate_CMS(int AgencyID, int Year)
        {
            int worksheet = 2;
            int row = 21;
            string sql;
            sql = string.Format(@"
SELECT Controls.Control
      ,FAMILY_ID
      ,FAMILY_NAME
      ,TYPE
      ,NAME
      ,DESCRIPTION
      ,IMPLEMENTATION_STANDARDS
      ,GUIDANCE
      ,ASSESSMENT_METHOD
      ,RELATED_TO
      ,IMPLEMENTATION
      ,EVIDENCE_DESCRIPTION
      ,EVIDENCE_DELIVERY
      ,EVIDENCE_FILES
      ,AUDIT_STATUS
      ,PREAUDIT_STATUS
      ,PREAUDIT_REVIEW_DATE
      ,EVIDENCE_SUBMIT_DATE
      ,STATUS
      ,CRITICALITY
      ,LAST_AUDIT
      ,(SELECT VALUE FROM REPORT_DETAILS WHERE NAME = 'Control::Information' AND VERSION = AUDIT_TYPE) version
      ,(SELECT VALUE FROM REPORT_DETAILS WHERE NAME = 'Control::Title' AND VERSION = {0}) Title
FROM CONTROLS
INNER JOIN AUDIT_CONTROLS
ON CONTROLS.CONTROL_ID = AUDIT_CONTROLS.CONTROL_ID
WHERE AUDIT = '{0}' AND AGENCY_ID = '{1}'
ORDER BY CONTROLS.CONTROL;"
                               ,Year
                               ,AgencyID);
            Require.False(null == mydatabase, "impossible error in Generate");
#pragma warning disable CS8602 // Dereference of a possibly null reference.
            DataTable temp = mydatabase.GetDataTable(sql);
#pragma warning restore CS8602 // Dereference of a possibly null reference.
            Require.False(null == Report, "impossible error in Generate 2");
#pragma warning disable CS8602 // Dereference of a possibly null reference.
            Report.Workbook.Worksheets.Add("Summary");
#pragma warning restore CS8602 // Dereference of a possibly null reference.

            foreach (DataRow i in temp.Rows)
            {
                object? hold = i["CONTROL"];
                Require.False(null == hold, "impossible error in Generate 3");
#pragma warning disable CS8602 // Dereference of a possibly null reference.
#pragma warning disable CS8600 // Converting null literal or possible null value to non-nullable type.
                string name = hold.ToString();
#pragma warning restore CS8600 // Converting null literal or possible null value to non-nullable type.
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                Report.Workbook.Worksheets.Add(name);
                CMS_Control_Page_Generator.CMS_Control_Page_Generator.Add_Control(Report.Workbook.Worksheets[name], i, fresh);
                worksheet += 1;
                CMS_Front_Page_Generator.CMS_Front_Page_Generator.Add_Frontpage_row(Report.Workbook.Worksheets["Summary"], i, row);
                row += 1;
            }

            // Get the audit Type
            sql = string.Format(@"
SELECT MAX(AUDIT_TYPE)
FROM CONTROLS
INNER JOIN AUDIT_CONTROLS
ON CONTROLS.CONTROL_ID = AUDIT_CONTROLS.CONTROL_ID
WHERE AUDIT = '{0}' AND AGENCY_ID = '{1}'
ORDER BY CONTROLS.CONTROL;"
                                ,Year
                                ,AgencyID);
            string version = mydatabase.ExecuteScalar(sql);

            // Get page Title
            sql = string.Format(@"
SELECT VALUE
FROM REPORT_DETAILS
WHERE NAME = 'Frontpage::Title'
  AND VERSION = {0}"
                               ,Year);
            string Title = mydatabase.ExecuteScalar(sql);

            // Get Agency info
            sql = string.Format(@"
SELECT AGENCY_ID
      ,TEAM_ID
      ,SYSTEM_ID
      ,PLATFORM
FROM AGENCIES
WHERE AGENCY_ID = '{0}';"
                               ,AgencyID);
            temp = mydatabase.GetDataTable(sql);
            CMS_Front_Page_Generator.CMS_Front_Page_Generator.Add_Frontpage_Header(Report.Workbook.Worksheets["Summary"], row - 1, temp, Title, version);
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
