using System.Data;
using common;
using OfficeOpenXml;

namespace CMS_Rename_Agency
{
    struct columns
    {
        public string Excel_Name;
        public string Database_Name;
        public int Column;
        public columns(string name, int column)
        {
            this.Excel_Name = name;
            this.Database_Name = name;
            this.Column = column;
        }
        public columns(string excel, string database, int column)
        {
            this.Excel_Name = excel;
            this.Database_Name = database;
            this.Column = column;
        }
    }
    public partial class Form1 : Form
    {
        readonly columns[] Excel_Columns = {
                                            new("AGENCY_ID", 1)
                                           ,new("TEAM_ID", 2)
                                           ,new("SYSTEM_ID", 3)
                                           ,new("PLATFORM", 4)
                                           ,new("NICK_ID", 5)
                                           ,new("CONTACT",6)
                                           ,new("FINAL_REPORT_KEY",7)
                                           ,new("PLATFORM_RESPONSE_PROVIDER",8)
                                           };
        string[]? files;
        static private SQLiteDatabase? mydatabase;
        public bool yes_to_all;
        public bool abort_all;
        public int change_count;
        public string Database = "audits.db";

        public Form1()
        {
            try
            {
                Require.True(File.Exists(Database), "unable to locate database file: " + Database, false);
            }
            catch (Exception ex) 
            { 
                MessageBox.Show(ex.Message);
                Environment.Exit(1);
            }

            InitializeComponent();
            this.AllowDrop = true;
#pragma warning disable CS8622 // Nullability of reference types in type of parameter doesn't match the target delegate (possibly because of nullability attributes).
            this.DragDrop += new DragEventHandler(UserControl1_DragDrop);
            this.DragEnter += new DragEventHandler(UserControl1_DragEnter);
            this.DragLeave += new EventHandler(UserControl1_DragLeave);
#pragma warning restore CS8622 // Nullability of reference types in type of parameter doesn't match the target delegate (possibly because of nullability attributes).
            // Target File we are updating

            yes_to_all = false;
            abort_all = false;
            change_count = 0;

            mydatabase = new SQLiteDatabase(Database);
        }

        void UserControl1_DragLeave(object sender, EventArgs e)
        {
        }

        void UserControl1_DragEnter(object sender, DragEventArgs e)
        {
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
            if (null == e) return;
            try
            {
                Require.False(null == e.Data, "Failed to receive any files");
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                files = (string[])e.Data.GetData(DataFormats.FileDrop);
#pragma warning restore CS8602 // Dereference of a possibly null reference.
            }
            catch (Exception ex) 
            {
                MessageBox.Show(ex.Message);
            }

            label1.Text = "Importing";
            this.BackColor = Color.Red;

            var are = new AutoResetEvent(false);
#pragma warning disable CS8622 // Nullability of reference types in type of parameter doesn't match the target delegate (possibly because of nullability attributes).
            bool v = ThreadPool.QueueUserWorkItem(RunMySlowLogicWrapped, are);
#pragma warning restore CS8622 // Nullability of reference types in type of parameter doesn't match the target delegate (possibly because of nullability attributes).
            if (v) are.WaitOne();

            if (abort_all)
            {
                label1.Text = "Agency import Aborted";
                this.BackColor = Color.Coral;
            }
            else if (0 != change_count)
            {
                label1.Text = string.Format("{0} Agencies records changed", change_count);
                this.BackColor = Color.Green;
            }
            else
            {
                label1.Text = "No changes detected";
                this.BackColor = Color.Green;

            }
        }

        private void RunMySlowLogicWrapped(Object state)
        {
            AutoResetEvent are = (AutoResetEvent)state;
            if (null == files)
            {
                are.Set();
                return;
            }
            foreach (string f in files)
            {
                Read_agency_book(f);
            }
            are.Set();
        }

        private void Read_agency_book(string FileName)
        {
            FileInfo name = new(FileName);
            Require.True(name.Exists, "Not a real file");
            ExcelPackage package = new(name);
            Require.False(null == package, "isn't a real excel file");
#pragma warning disable CS8602 // Dereference of a possibly null reference.
            ExcelWorkbook book = package.Workbook;
#pragma warning restore CS8602 // Dereference of a possibly null reference.
            Require.False(null == book, "Excel file has no workbooks");

#pragma warning disable CS8602 // Dereference of a possibly null reference.
            foreach (ExcelWorksheet sheet in book.Worksheets)
            {
                Read_agency_page(sheet);
            }
#pragma warning restore CS8602 // Dereference of a possibly null reference.
        }

        private void Read_agency_page(ExcelWorksheet ws)
        {
            if (null == mydatabase) return;
            string sql;
            int AgencyID;
            string fields = string.Join(", ", Excel_Columns.Select(x => x.Database_Name));
            DataTable dt;

            for (int i = 2; "" != ws.Cells[i, 1].Text; i += 1)
            {
                try
                {
                    AgencyID = Convert.ToInt32(ws.Cells[i, 1].Value);
                    sql = string.Format("SELECT {0} FROM AGENCIES WHERE AGENCY_ID = {1}", fields, AgencyID);
                    dt = mydatabase.GetDataTable(sql);
                    if (0 == dt.Rows.Count)
                    {
                        sql = string.Format("INSERT INTO AGENCIES (AGENCY_ID) VALUES ('{0}');"
                                           ,AgencyID.ToString());
                        mydatabase.ExecuteScalar(sql);
                    }
                    foreach (DataRow r in dt.AsEnumerable())
                    {
                        foreach (columns c in Excel_Columns)
                        {
                            if (abort_all) return;
#pragma warning disable CS8604 // Possible null reference argument.
                            string hold = get_cell_value(ws, i, c.Column, c.Excel_Name);
                            if (!Match.CaseInsensitive("NULL", hold)) Update_row(AgencyID, hold, Convert.ToString(r[c.Database_Name]), c.Database_Name);
                            else Console.WriteLine(string.Format("did not find {0} in column {1}", c.Column, c.Column));
#pragma warning restore CS8604 // Possible null reference argument.
                            if (abort_all) return;
                        }
                    }
                }
                catch
                {
                    MessageBox.Show(string.Format("Line {0} contains content which is not valid", i));
                }
            }
        }

        private string get_cell_value(ExcelWorksheet ws, int row, int column, string header_name)
        {
            string header = ws.Cells[1, column].Text;
            if (!Match.CaseInsensitive(header, header_name)) return "NULL";

            try
            {
                string r = ws.Cells[row, column].Text;
                r = r.Trim();
                return r;
            }
            catch
            {
                return "";
            }
        }
        private void Update_row(int AgencyID, string s1, string s2, string column_name)
        {
            /* Nothing to do if matching  (Always case sensitive) */
            if (Match.CaseSensitive(s1, s2)) return;

            change_count += 1;
            Change_row(AgencyID, s1, column_name);
        }

        private static void Change_row(int AgencyID, string s, string column_name)
        {
            if(null == mydatabase) return;
            string sql = string.Format("UPDATE AGENCIES SET {0} = '{1}' WHERE AGENCY_ID = {2};", column_name, s.Replace("'", "''"), AgencyID);
            mydatabase.ExecuteScalar(sql);
        }
    }
}
