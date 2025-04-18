using common;

namespace CMS_Controls_Export
{
    internal class Program
    {
        static void Main(string[] args)
        {
            // Where we are going to put our output
            string filename = "Results " + DateTime.Now.ToString("MM_dd_yyyy__hh_mm_ss") + ".xlsx";
            // What Audit Number to use
            int Audit_number = 0;
            string Audit_Name = string.Empty;
            // Target File we are updating
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
                else if (Match.CaseInsensitive("--audit-name", args[i]))
                {
                    Audit_Name = args[i + 1];
                    Require.True(null != Audit_Name, "--audit-name requires an argument\n");
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                    Require.False(Audit_Name.Equals(""), "--audit-name can't use an empty string\n");
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                    i += 2;
                }
                else if (Match.CaseInsensitive("--audit-number", args[i]))
                {
                    string hold = args[i + 1];
                    Require.True(null != hold, "--audit-number requires an argument\n");
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                    Require.False(hold.Equals(""), "--audit-number can't use an empty string\n");
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                    try
                    {
                        Audit_number = Convert.ToInt32(hold);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex);
                        Console.WriteLine("Audit number (--audit-number) must be a number");
                        Environment.Exit(1);
                    }
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

            FileInfo output = new FileInfo(filename);

            if ((0 == Audit_number) && Match.CaseInsensitive(Audit_Name, string.Empty))
            {
                Console.WriteLine("Audit number (--audit-number ####) or Audit name (--audit-name \"%%%%\") must be provided");
                Environment.Exit(1);
            }

            if (Match.CaseInsensitive(Database, string.Empty))
            {
                Console.WriteLine("Database name (--database-name %%%%) must be provided");
                Environment.Exit(1);
            }

            if (verbose)
            {
                Console.WriteLine(string.Format("starting generation of: {0}", filename));
            }

            CMS_Controls_Export gen = new CMS_Controls_Export();
            if (0 == Audit_number) Audit_number = gen.Initialize(Database, Audit_Name);
            else Audit_Name = gen.Initialize(Database, Audit_number);
            gen.Generate(output, Audit_number, Audit_Name);
            gen.MultiSave();

            if (verbose)
            {
                Console.WriteLine(string.Format("generation of: {0} is complete", filename));
            }
        }
    }
}