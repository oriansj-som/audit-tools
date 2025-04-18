using common;
using OfficeOpenXml;
using System.Data.SQLite;

namespace Excel_Importer
{
    internal class Excel_File_Importer
    {
        private SQLiteDatabase mydatabase;
        private ExcelWorkbook book;
        private bool verbose;

        public void Initialize(string Database, FileInfo existingFile, bool v)
        {
            verbose = v;
            if (!File.Exists(Database))
            {
                if (verbose)
                {
                    Console.WriteLine(string.Format("Database file {0} does not exist, creating", Database));
                }

                SQLiteConnection.CreateFile(Database);
            }

            // Setup for reading file
            book = new ExcelPackage(existingFile).Workbook;
            // Connect to Database
            mydatabase = new SQLiteDatabase(Database);
        }

        public void import()
        {
            foreach(ExcelWorksheet ws in book.Worksheets)
            {
                string name = ws.Name;
                Excel_Tab_Importer tab = new Excel_Tab_Importer();
                string sql = tab.initialize(name, ws, verbose);
                mydatabase.ExecuteNonQuery(sql);
                List<Dictionary<String, String>> v = tab.import();
                int row = 2;
                mydatabase.ExecuteNonQuery("BEGIN TRANSACTION");
                foreach (Dictionary<String, String> i in v)
                {
                    mydatabase.Insert(name, i);
                    if(verbose)
                    {
                        Console.WriteLine(string.Format("imported up to row: {0}", row));
                    }
                    row += 1;
                }
                mydatabase.ExecuteNonQuery("COMMIT TRANSACTION");
            }
        }
    }
}
