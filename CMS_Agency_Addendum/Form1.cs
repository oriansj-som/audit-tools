using common;
using System.Data;

namespace CMS_Agency_Addendum
{
    public partial class Form1 : Form
    {
        private static SQLiteDatabase? mydatabase;
        private static readonly string Database = "audits.db";
        private string[]? files;
        private static int Year;

        public Form1()
        {
            InitializeComponent();
            this.AllowDrop = true;
#pragma warning disable CS8622 // Nullability of reference types in type of parameter doesn't match the target delegate (possibly because of nullability attributes).
            this.DragDrop += new DragEventHandler(UserControl1_DragDrop);
            this.DragEnter += new DragEventHandler(UserControl1_DragEnter);
            this.DragLeave += new EventHandler(UserControl1_DragLeave);
#pragma warning restore CS8622 // Nullability of reference types in type of parameter doesn't match the target delegate (possibly because of nullability attributes).

            Year = DateTime.Today.Year;
            try
            {
                Require.True(File.Exists(Database), "Unable to open database: " + Database, false);
                Require.True(File.Exists(".\\CMS_Subset_Update.exe"), "Required subtool CMS_Subset_Update.exe not found", false);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                Environment.Exit(3);
            }

            // Connect to Database
            mydatabase = new SQLiteDatabase(Database);
            try
            {
                Require.False(null == mydatabase, string.Format("File {0} is not a valid sqlite database", Database), false);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                Environment.Exit(3);
            }

#pragma warning disable CS8602 // Dereference of a possibly null reference.
            DataTable versions = mydatabase.GetDataTable("SELECT NAME FROM AUDIT_TYPES ORDER BY AUDIT_TYPE DESC;");
#pragma warning restore CS8602 // Dereference of a possibly null reference.
            foreach (DataRow i in versions.Rows)
            {
                comboBox1.Items.Add(i["NAME"].ToString());
            }

            // Default select current year
            comboBox1.SelectedIndex= 0;
        }

        void UserControl1_DragLeave(object sender, EventArgs e)
        {
        }

        void UserControl1_DragEnter(object sender, DragEventArgs e)
        {
            // silence compiler warning
            if (null == e) return;
            if (null == e.Data) return;

            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Copy;
            }
            else
            {
                e.Effect = DragDropEffects.None;
            }
        }
        void UserControl1_DragDrop(object sender, DragEventArgs e)
        {
            // silence compiler warning
            if (null == e) return;
            if (null == e.Data) return;

            files = (string[])e.Data.GetData(DataFormats.FileDrop);

            label1.Text = "Exporting";
            this.BackColor = Color.Red;

            foreach (string f in files)
            {
                Generate_Addendum(f);
            }

            label1.Text = "Controls exported";
            this.BackColor = Color.Green;
        }

        static void Generate_Addendum(string FileName)
        {
            Populate_Control_Assignments assignments = new();
            FileInfo existingFile = new(FileName);
            List<assignments> set = assignments.Process(Database, existingFile, "CMS ARC-AMPE");
            assignments.Cleanup();

            List<string> agencies = set.Select(x => x.agency_nick).Distinct().ToList();
            foreach (string agency in agencies)
            {
                string controls = "'" + string.Join("', '", set.Where(x => Match.CaseInsensitive(agency, x.agency_nick)).Select(x => x.control_name)) + "'";
                controls = controls.Replace("''", "").Trim();
                if (Match.CaseInsensitive(string.Empty, controls)) continue;

                string args1 = string.Format("--audit {0} --database-name {1} --agency-name \"{2}\" --controls \"{3}\" --verbose"
                                            ,Year.ToString()
                                            ,Database
                                            ,agency
                                            ,controls.Replace("'", "").Replace(" ", ""));
                var process1 = System.Diagnostics.Process.Start(".\\CMS_Subset_Update.exe", args1);
                process1.WaitForExit();

                if (0 != process1.ExitCode)
                {
                    MessageBox.Show("Unknown Entry: " + agency + " :: " + controls);
                    return;
                }

                string output = Year.ToString() + " Audit " + agency + " - addendum.xlsx";
                // To keep things simple all all of the binaries in a single folder
                // And put the output in a parent directory
                string args2 = string.Format("--audit {0} --agency-name \"{1}\" --database-name {2} --output-file \"{3}\" --controls \"{4}\" --verbose"
                                            ,Year.ToString()
                                            ,agency
                                            ,Database
                                            ,output 
                                            ,controls);
                var process2 = System.Diagnostics.Process.Start(".\\CMS_Subset_Export.exe", args2);
                process2.WaitForExit();
            }
        }
    }
}
