using common;

namespace CMS_Update_Audit_Controls
{
    internal class Program
    {
        static void Main(string[] args)
        {
            // Get the file we are going to process
            FileInfo? existingFile = null;
            // Target File we are updating
            string Database = String.Empty;
            // Are we doing a special audit?
            string Audit_Name = string.Empty;
            bool verbose = false;

            int i = 0;
            while (i < args.Length)
            {
                if (Match.CaseInsensitive("--input-file", args[i]) || Match.CaseInsensitive("-i", args[i]))
                {
                    string filename = args[i + 1];
                    Require.True(null != filename, "--input-file requires an argument\n");
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                    Require.False(filename.Equals(""), "--input-file can't use an empty string\n");
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                    existingFile = new FileInfo(filename);
                    Require.True(existingFile.Exists, String.Format("Unable to find {0}", existingFile));
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
                else if (Match.CaseInsensitive("--audit-name", args[i]) || Match.CaseInsensitive("-A", args[i]))
                {
                    Audit_Name = args[i + 1];
                    Require.True(null != Audit_Name, "--audit-name requires an argument\n");
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                    Require.False(Audit_Name.Equals(""), "--audit-name can't use an empty string\n");
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
                    Environment.Exit(1);
                }
            }

            if (null == existingFile)
            {
                Console.WriteLine("input file (--input-file %%%%) is required");
                Environment.Exit(1);
            }

            if (Match.CaseInsensitive(Database, string.Empty))
            {
                Console.WriteLine("Database file (--database-name %%%%) is required");
                Environment.Exit(1);
            }

            if (verbose)
            {
                Console.WriteLine("Beginning import");
            }

            Populate_Control_Assignments assignments= new();
            assignments.Process(Database, existingFile, Audit_Name);
            assignments.Cleanup();

            if (verbose)
            {
                Console.WriteLine("Import Complete");
            }
        }
    }
}