using common;

namespace Excel_Importer
{
    internal class Program
    {
        static void Main(string[] args)
        {
            // Where we are reading from
            string filename = string.Empty;
            // Where our data is
            string Database = string.Empty;
            // How chatty this app should be
            bool verbose = false;

            int i = 0;
            while (i < args.Length)
            {
                if (Match.CaseInsensitive("--input-file", args[i]) || Match.CaseInsensitive("-i", args[i]))
                {
                    filename = args[i + 1];
                    Require.True(null != filename, "--input-file requires an argument\n");
                    Require.False(filename.Equals(""), "--input-file can't use an empty string\n");
                    i += 2;
                }
                else if (Match.CaseInsensitive("--database-name", args[i]) || Match.CaseInsensitive("-d", args[i]))
                {
                    Database = args[i + 1];
                    Require.True(null != Database, "--database-name requires an argument\n");
                    Require.False(Database.Equals(""), "--database-name can't use an empty string\n");
                    i += 2;
                }
                else if (Match.CaseInsensitive("--help", args[i]) || Match.CaseInsensitive("-h", args[i]))
                {
                    Console.WriteLine("Excel_Importer is an Excel File to SQLite Database table import tool\n"
                                     +"Each tab in your excel file will become a separate table\n"
                                     +"Each column_name will become a column name in that sheet's corresponding table\n"
                                     +"Each column can also contain [type] information\n"
                                     +"In the Column_Name [type] format, eg:\n"
                                     +"ID [INTEGER NOT NULL PRIMARY KEY AUTOINCREMENT UNIQUE]\n"
                                     +"to identify what database file to use (or create) --database-name filename\n"
                                     +"to identify what excel file to load use --input-file filename");
                    Environment.Exit(0);
                }
                else if (Match.CaseInsensitive("--verbose", args[i]) || Match.CaseInsensitive("-v", args[i]))
                {
                    int index = 0;
                    verbose = true;
                    foreach (string s in args)
                    {
                        Console.WriteLine(string.Format("argument {0}: {1}", index, s));
                        index = index + 1;
                    }
                    i += 1;
                }
                else
                {
                    Console.WriteLine(string.Format("Unknown argument: {0} received", args[i]));
                    i += 1;
                }
            }

            if(Database.Equals(string.Empty))
            {
                Console.WriteLine("Database name must be provided (--database-name filename)");
                Environment.Exit(1);
            }

            if(filename.Equals(string.Empty))
            {
                Console.WriteLine("An input file name must be provided (--input-file filename)");
                Environment.Exit(1);
            }

            if (verbose) Console.WriteLine("Starting file load");
            Excel_File_Importer f = new Excel_File_Importer();
            f.Initialize(Database, new FileInfo(filename), verbose);
            f.import();
            if (verbose) Console.WriteLine("processing complete");
        }
    }
}