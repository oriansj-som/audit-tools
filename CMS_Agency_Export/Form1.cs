using System.Data;
using common;

namespace CMS_Agency_Export
{
    public partial class Form1 : Form
    {
        static private SQLiteDatabase? mydatabase;
        static private readonly string Database = "audits.db";
        static private readonly string prepath = "..\\";

        // For passing to long running programs
        public class Subprocess
        {
            public string? program;
            public string? args;
            public string? message;
            public int return_value;
        }
        static private List<Subprocess>? work;

        public Form1()
        {
            InitializeComponent();

            try
            {
                Require.True(File.Exists(Database), "Unable to open database: " + Database, false);
                Require.True(File.Exists(".\\CMS_Audit_Export.exe"), "Unable to open required tool: .\\CMS_Audit_Export.exe", false);
                Require.True(File.Exists(".\\CMS_Year_Generator.exe"), "Unable to open required tool: .\\CMS_Year_Generator.exe", false);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                Environment.Exit(3);
            }

            // Connect to Database
            mydatabase = new SQLiteDatabase(Database);
            DataTable Years;
            try
            {
                Require.False(null == mydatabase, string.Format("File {0} is not a valid sqlite database", Database), false);
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                Years = mydatabase.GetDataTable("SELECT DISTINCT AUDIT FROM AUDIT_CONTROLS ORDER BY AUDIT DESC;");
#pragma warning restore CS8602 // Dereference of a possibly null reference.


                comboBox1.Items.Add("NEW");
                foreach (DataRow i in Years.Rows)
                {
                    comboBox1.Items.Add(i[0].ToString());
                }

                DataTable Agencies = mydatabase.GetDataTable("SELECT TEAM_ID, SYSTEM_ID, PLATFORM, NICK_ID FROM AGENCIES;");
                comboBox2.Items.Add("ALL");
                foreach (DataRow i in Agencies.Rows)
                {
#pragma warning disable CS8604 // Possible null reference argument.
                    if (Match.CaseInsensitive("RETIRED", i["NICK_ID"].ToString()))
                    {
                        if (Match.CaseInsensitive("System", i["SYSTEM_ID"].ToString()))
                        {
                            comboBox2.Items.Add(i["TEAM_ID"].ToString() + " - RETIRED");
                        }
                        else
                        {
                            comboBox2.Items.Add(i["SYSTEM_ID"].ToString() + " - RETIRED");
                        }
                    }
                    else if (Match.CaseInsensitive("", i["NICK_ID"].ToString()))
                    {
                        comboBox2.Items.Add(i["TEAM_ID"].ToString() + " " + i["SYSTEM_ID"].ToString() + " " + i["PLATFORM"].ToString());
                    }
                    else
                    {
                        comboBox2.Items.Add(i["NICK_ID"].ToString());
                    }
#pragma warning restore CS8604 // Possible null reference argument.
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                Environment.Exit(3);
            }

        }

        private void Generate_report_Click(object sender, EventArgs e)
        {
            int Year;
            string Agency;
            int AgencyID;

            if (null == comboBox1.SelectedItem)
            {
                comboBox1.SelectedIndex = comboBox1.FindStringExact(comboBox1.Text);
                if (-1 == comboBox1.SelectedIndex)
                {
                    MessageBox.Show("Please Select a year");
                    return;
                }
            }

#pragma warning disable CS8602 // Dereference of a possibly null reference.
#pragma warning disable CS8604 // Possible null reference argument.
            if (Match.CaseInsensitive(comboBox1.SelectedItem.ToString(), "NEW"))
            {
                string extra = "";
                if(checkBox1.Checked)
                {
                    extra = " --upgrade";
                }
                Year = DateTime.Today.Year;
                string args = string.Format("--audit-year {0} --database-name {1}{2}"
                                           , Year.ToString()
                                           ,Database
                                           ,extra);
                var process = System.Diagnostics.Process.Start(".\\CMS_Year_Generator.exe", args);
                process.WaitForExit();
                if (process.ExitCode != 0) MessageBox.Show("CMS_Year_Generator.exe failed to run");
            }
#pragma warning restore CS8604 // Possible null reference argument.
#pragma warning restore CS8602 // Dereference of a possibly null reference.
            else
            {
                try
                {
                    // What Audit Number to use
#pragma warning disable CS8600 // Converting null literal or possible null value to non-nullable type.
                    string ID = comboBox1.SelectedItem.ToString();
#pragma warning restore CS8600 // Converting null literal or possible null value to non-nullable type.
                    Year = Convert.ToInt16(ID);
                }
                catch
                {
                    MessageBox.Show("Please Select a year");
                    return;
                }
            }

            if (null == comboBox2.SelectedItem)
            {
                comboBox2.SelectedIndex = comboBox2.FindStringExact(comboBox2.Text);
                if (-1 == comboBox2.SelectedIndex)
                {
                    MessageBox.Show("Please Select an Agency or ALL");
                    return;
                }
            }

            try
            {
#pragma warning disable CS8602 // Dereference of a possibly null reference.
#pragma warning disable CS8600 // Converting null literal or possible null value to non-nullable type.
                Agency = comboBox2.SelectedItem.ToString();
#pragma warning restore CS8600 // Converting null literal or possible null value to non-nullable type.
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                AgencyID = comboBox2.SelectedIndex;
            }
            catch
            {
                MessageBox.Show("Please Select an Agency or ALL");
                return;
            }

            if (null == Agency) return;
            work = new List<Subprocess>();
            if (Match.CaseInsensitive("ALL", Agency))
            {
                Print_all(Year);
            }
            else
            {
                work.Add(Print_Agency(Year, Agency, AgencyID));
                Run_work();
                MessageBox.Show("The file: " + Year.ToString() + " Audit " + Agency + ".xlsx contains your report");
            }
        }

        private static Subprocess Print_Agency(int Year, string Agency, int AgencyID)
        {
            string output = prepath + Year.ToString() + " Audit " + Agency + ".xlsx";
            // To keep things simple all all of the binaries in a single folder
            // And put the output in a parent directory
            string args = string.Format("--audit {0} --agency-number {1} --database-name {2} --output-file \"{3}\" --verbose"
                                       , Year.ToString()
                                       , AgencyID.ToString()
                                       , Database
                                       , output);

            // Queue up the work
            Subprocess s = new()
            {
                args = args,
                program = ".\\CMS_Audit_Export.exe",
                message = output + " generation failed"
            };
            return s;

        }
        private void Run_work()
        { 

            // Run the job
            var are = new AutoResetEvent(false);
#pragma warning disable CS8622 // Nullability of reference types in type of parameter doesn't match the target delegate (possibly because of nullability attributes).
            bool v = ThreadPool.QueueUserWorkItem(RunMySlowLogicWrapped, are);
#pragma warning restore CS8622 // Nullability of reference types in type of parameter doesn't match the target delegate (possibly because of nullability attributes).
            if (v) are.WaitOne();

            // Sanity check for failures
            if (null == work) return;
            foreach (Subprocess i in work)
            {
                if(0 != i.return_value)
                {
                    MessageBox.Show(i.message);
                }
            }
        }

        void Print_all(int Year)
        {
            if (null == mydatabase) return;
            if (null == work) return;
            try 
            { 
                // Get everything we are collecting
                string sql = string.Format("SELECT DISTINCT AGENCIES.AGENCY_ID, TEAM_ID, SYSTEM_ID, PLATFORM, NICK_ID FROM AGENCIES INNER JOIN AUDIT_CONTROLS ON AGENCIES.AGENCY_ID = AUDIT_CONTROLS.AGENCY_ID WHERE AUDIT = '{0}';"
                                          ,Year);
                // Queue up the questionaires
                DataTable Agencies = mydatabase.GetDataTable(sql);
                foreach (DataRow i in Agencies.Rows)
                {
                    string Agency;
#pragma warning disable CS8604 // Possible null reference argument.
                    int AgencyID = int.Parse(i["AGENCY_ID"].ToString());
                    if (Match.CaseInsensitive("RETIRED", i["NICK_ID"].ToString()))
                    {
                        // Skip Retired Systems
                        continue;
                    }
                    else if (Match.CaseInsensitive("", i["NICK_ID"].ToString()))
                    {
                        Agency = i["TEAM_ID"].ToString() + " " + i["SYSTEM_ID"].ToString() + " " + i["PLATFORM"].ToString();
                    }
                    else
                    {
#pragma warning disable CS8600 // Converting null literal or possible null value to non-nullable type.
                        Agency = i["NICK_ID"].ToString();
#pragma warning restore CS8600 // Converting null literal or possible null value to non-nullable type.
                    }
#pragma warning restore CS8604 // Possible null reference argument.

                    Require.False(null == work, "Failed to initialize work", false);
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                    if (null == Agency) continue;
                    work.Add(Print_Agency(Year, Agency, AgencyID));
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                }

                // Process them all
                Run_work();
                }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                Environment.Exit(3);
            }
            MessageBox.Show("All Agencies for " + Year.ToString() + " have been exported");
        }

        private void RunMySlowLogicWrapped(Object state)
        {
            AutoResetEvent are = (AutoResetEvent)state;
            if (null == work)
            {
                are.Set();
                return;
            }

            // Go hard
            ParallelOptions po = new()
            {
                MaxDegreeOfParallelism = Environment.ProcessorCount << 4
            };

            Parallel.ForEach(work, po, i =>
            {
                Require.False(null == i.program, "No program provided", false);
                Require.False(null == i.args, "No arguments provided", false);
#pragma warning disable CS8604 // Possible null reference argument.
                System.Diagnostics.Process process = System.Diagnostics.Process.Start(i.program, i.args);
#pragma warning restore CS8604 // Possible null reference argument.
                process.WaitForExit();
                i.return_value = process.ExitCode;
            });
            are.Set();
            return;
        }
    }
}

