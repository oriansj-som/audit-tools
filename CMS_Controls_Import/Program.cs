using common;

namespace CMS_Controls_Import
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string filename = string.Empty;
            string Database = string.Empty;
            string Audit_Name = string.Empty;
            bool verbose = false;

            int i = 0;
            while (i < args.Length)
            {
                if (Match.CaseInsensitive("--input-file", args[i]) || Match.CaseInsensitive("-i", args[i]))
                {
                    filename = args[i + 1];
                    Require.True(null != filename, "--input-file requires an argument\n");
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                    Require.False(filename.Equals(""), "--input-file can't use an empty string\n");
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                    i += 2;
                }
                else if (Match.CaseInsensitive("--audit-name", args[i]) || Match.CaseInsensitive("-a", args[i]))
                {
                    Audit_Name = args[i + 1];
                    Require.True(null != Audit_Name, "--audit-name requires an argument\n");
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                    Require.False(Audit_Name.Equals(""), "--audit-name can't use an empty string\n");
#pragma warning restore CS8602 // Dereference of a possibly null reference.
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

            Require.False(Match.CaseInsensitive(string.Empty, filename), "input file name (--input-file %%%%) must be provided");
            FileInfo input = new FileInfo(filename);

            if(!input.Exists)
            {
                Console.WriteLine("input file name (--input-file %%%%) must exist");
                Environment.Exit(1);
            }

            if(Match.CaseInsensitive(string.Empty, Audit_Name))
            {
                Console.WriteLine("Audit name (--audit-name \"%%%%\") must be provided");
                Environment.Exit(1);
            }

            if (Match.CaseInsensitive(Database, string.Empty))
            {
                Console.WriteLine("Database name (--database-name %%%%) must be provided");
                Environment.Exit(1);
            }

            if (verbose)
            {
                Console.WriteLine(string.Format("starting import  of: {0}", filename));
            }

            CMS_Controls_Import control = new CMS_Controls_Import();
            string auditType = control.Initialize(Database, input, Audit_Name);
            control.Import(auditType, Audit_Name);
            control.Cleanup();

            if (verbose)
            {
                Console.WriteLine(string.Format("import of: {0} is complete", filename));
            }
        }
    }
}