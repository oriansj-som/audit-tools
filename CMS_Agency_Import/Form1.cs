using common;

namespace CMS_Agency_Import
{
    public partial class Form1 : Form
    {
        private readonly string Database = "audits.db";
        private string[]? files;
        public string? final_message;
        int count = 0;
        public Form1()
        {
            InitializeComponent();
            this.AllowDrop = true;
#pragma warning disable CS8622 // Nullability of reference types in type of parameter doesn't match the target delegate (possibly because of nullability attributes).
            this.DragDrop += new DragEventHandler(UserControl1_DragDrop);
            this.DragEnter += new DragEventHandler(UserControl1_DragEnter);
            this.DragLeave += new EventHandler(UserControl1_DragLeave);
#pragma warning restore CS8622 // Nullability of reference types in type of parameter doesn't match the target delegate (possibly because of nullability attributes).

            try
            {
                Require.True(File.Exists(Database), "Unable to open database: " + Database, false);
                Require.True(File.Exists(".\\CMS_Audit_Import.exe"), "Subtool: CMS_Audit_Import.exe not found", false);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                Environment.Exit(3);
            }
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
            if (null == e.Data) return;
            files = (string[])e.Data.GetData(DataFormats.FileDrop);

            label1.Text = "Importing";
            this.BackColor = Color.Red;
            final_message = string.Empty;
            try
            {
                var are = new AutoResetEvent(false);
#pragma warning disable CS8622 // Nullability of reference types in type of parameter doesn't match the target delegate (possibly because of nullability attributes).
                bool v = ThreadPool.QueueUserWorkItem(RunMySlowLogicWrapped, are);
#pragma warning restore CS8622 // Nullability of reference types in type of parameter doesn't match the target delegate (possibly because of nullability attributes).
                if (v) are.WaitOne();
            }
            catch
            {
                label1.Text = "Unexpected issue with input files";
                this.BackColor = Color.Tomato;
                return;
            }

            label1.Text = final_message;
            if (final_message.Equals("Files imported", StringComparison.CurrentCultureIgnoreCase))
            {
                label1.Text = string.Format("{0} Files successfully imported", count);
                this.BackColor = Color.Green;
            }
            else
            {
                this.BackColor = Color.Tomato;
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

            foreach (string i in files)
            {
                string args = String.Format("--database-name .\\audits.db --input-file \"{0}\"  --verbose", i);
                var process = System.Diagnostics.Process.Start(".\\CMS_Audit_Import.exe", args);
                process.WaitForExit();
                if (0 != process.ExitCode)
                {
                    final_message = "Failure to read " + i;
                    are.Set();
                    return;
                    
                }
                count += 1;
            }
            final_message = "Files imported";
            are.Set();
            return;
        }
    }
}