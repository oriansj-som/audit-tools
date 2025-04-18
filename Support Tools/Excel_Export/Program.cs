using common;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using System.Data;

namespace Excel_Export
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string filename = "dump " + DateTime.Now.ToString("MM_dd_yyyy__hh_mm_ss") + ".xlsx";
            string table = string.Empty;
            string Database = string.Empty;
            bool verbose = false;

            int i = 0;
            while (i < args.Length)
            {
                if (Match.CaseInsensitive("--output-file", args[i]) || Match.CaseInsensitive("-o", args[i]))
                {
                    filename = args[i + 1];
                    Require.True(null != filename, "--output-file requires an argument\n");
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                    Require.False(filename.Equals(""), "--output-file can't use an empty string\n");
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                    i += 2;
                }
                else if (Match.CaseInsensitive("--table", args[i]) || Match.CaseInsensitive("-t", args[i]))
                {
                    string hold = args[i + 1];
                    Require.True(null != hold, "--table requires an argument\n");
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                    Require.False(hold.Equals(""), "--table can't use an empty string\n");
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                    table = hold;
                    i += 2;
                }
                else if (Match.CaseInsensitive("--database-name", args[i]) || Match.CaseInsensitive("-d", args[i]))
                {
                    Database = args[i + 1];
                    Require.True(null != Database, "--database-name requires an argument\n");
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                    Require.False(Database.Equals(""), "--database-name can't use an empty string\n");
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                    i += 2;
                }
                else if (Match.CaseInsensitive("--verbose", args[i]))
                {
                    int index = 0;
                    verbose = true;
                    foreach (string s in args)
                    {
                        Console.WriteLine(string.Format("argument {0}: {1}", index, s));
                        index += 1;
                    }
                    i += 1;
                }
                else
                {
                    Console.WriteLine(string.Format("Unknown argument: {0} received", args[i]));
                    i += 1;
                }
            }

            if(string.IsNullOrEmpty(table))
            {
                Console.WriteLine("A table name must be provided");
                Environment.Exit(1);
            }

            if (string.IsNullOrEmpty(Database))
            {
                Console.WriteLine("A database name must be provided");
                Environment.Exit(1);
            }

            if (!File.Exists(Database))
            {
                Console.WriteLine("Database file does not exist");
                Environment.Exit(1);
            }

            if (verbose)
            {
                Console.WriteLine(string.Format("starting generation of: {0}", filename));
            }

            try
            {
                ExcelPackage? Report = new ExcelPackage();
                Report.Workbook.Worksheets.Add(table);
                ExcelWorksheet ws = Report.Workbook.Worksheets[table];

                SQLiteDatabase db = new SQLiteDatabase(Database);
                string sql = string.Format("SELECT * FROM {0};", table);
                DataTable dt = db.GetDataTable(sql);
                int column = 1;
                foreach (DataColumn dc in dt.Columns)
                {
                    ws.Cells[1, column].Value = dc.ColumnName;
                    int row = 2;
                    foreach (DataRow dr in dt.Rows)
                    {
                        ws.Cells[row, column].Value = dr[dc.ColumnName].ToString();
                        row += 1;
                    }
                    column += 1;
                }
                MultiSave(Report, new FileInfo(filename));
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Environment.Exit(1);
            }

            if (verbose)
            {
                Console.WriteLine(string.Format("generation of: {0} is complete", filename));
            }

        }

        // Enable us to incrementally save the workbook
        public static void MultiSave(ExcelPackage? Report, FileInfo? output)
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
