using common;

namespace CMS_Update_Agency_Controls
{
    public partial class Form1 : Form
    {
        private string[]? files;
        public List<string>? unknowns;
        private readonly string Database = "audits.db";

        public Form1()
        {
            try
            {
                Require.True(File.Exists(Database), "Unable to open database: " + Database, false);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                Environment.Exit(3);
            }
            InitializeComponent();
            this.AllowDrop = true;
#pragma warning disable CS8622 // Nullability of reference types in type of parameter doesn't match the target delegate (possibly because of nullability attributes).
            this.DragDrop += new DragEventHandler(UserControl1_DragDrop);
            this.DragEnter += new DragEventHandler(UserControl1_DragEnter);
            this.DragLeave += new EventHandler(UserControl1_DragLeave);
#pragma warning restore CS8622 // Nullability of reference types in type of parameter doesn't match the target delegate (possibly because of nullability attributes).
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

            var are = new AutoResetEvent(false);
            unknowns = new List<string>();
#pragma warning disable CS8622 // Nullability of reference types in type of parameter doesn't match the target delegate (possibly because of nullability attributes).
            ThreadPool.QueueUserWorkItem(RunMySlowLogicWrapped, are);
#pragma warning restore CS8622 // Nullability of reference types in type of parameter doesn't match the target delegate (possibly because of nullability attributes).
            are.WaitOne();

            label1.Text = "Controls imported";
            this.BackColor = Color.Green;
            if (0 != unknowns.Count)
            {
                MessageBox.Show("The following unknown Agencies were found: " + String.Join(", ", unknowns.ToArray()));
            }
        }

        private void RunMySlowLogicWrapped(Object state)
        {
            if (null == files) return;
            AutoResetEvent are = (AutoResetEvent)state;
            Import_Agency_Controls(files[0]);
            are.Set();
        }

        void Import_Agency_Controls(string FileName)
        {
            Populate_Control_Assignments assignments = new();
            FileInfo existingFile = new(FileName);
            assignments.Process(Database, existingFile, "CMS ARC-AMPE");
            unknowns = assignments.unknowns;
            assignments.Cleanup();
        }
    }
}
