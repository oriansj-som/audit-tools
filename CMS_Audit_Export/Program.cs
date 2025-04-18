using common;

namespace CMS_Audit_Export
{
    class Program
    {
        static void Main(string[] args)
        {
            // Where we are going to put our output
            string filename = "Results " + DateTime.Now.ToString("MM_dd_yyyy__hh_mm_ss") + ".xlsx";
            // What Audit Number to use
            int Audit = 0;
            //What Audit Type to use
            string Audit_Name = "CMS ARC-AMPE";
            // What Agency
            string Agency = string.Empty;
            int AgencyID = 0;
            // Target File we are updating
            string Database = string.Empty;
            bool verbose = false;
            bool fresh = true;

            int i = 0;
            while (i < args.Length)
            {
                if (Match.CaseInsensitive("--output-file", args[i]) || Match.CaseInsensitive("-o", args[i]))
                {
                    filename = args[i + 1];
                    Require.True(null != filename, "--output-file requires an argument\n");
                    Require.False(filename.Equals(""), "--output-file can't use an empty string\n");
                    i += 2;
                }
                else if (Match.CaseInsensitive("--audit", args[i]) || Match.CaseInsensitive("-a", args[i]))
                {
                    string hold = args[i + 1];
                    Require.True(null != hold, "--audit requires an argument\n");
                    Require.False(hold.Equals(""), "--audit can't use an empty string\n");
                    try
                    {
                        Audit = Convert.ToInt32(hold);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex);
                        Console.WriteLine("Audit number (--audit) must be a number");
                        Environment.Exit(1);
                    }
                    i += 2;
                }
                else if (Match.CaseInsensitive("--audit-type", args[i]) || Match.CaseInsensitive("-t", args[i]))
                {
                    string hold = args[i + 1];
                    Require.True(null != hold, "--audit requires an argument\n");
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                    Require.False(hold.Equals(""), "--audit can't use an empty string\n");
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                    Audit_Name = hold;
                    i += 2;
                }
                else if (Match.CaseInsensitive("--agency-number", args[i]) || Match.CaseInsensitive("-N", args[i]))
                {
                    string hold = args[i + 1];
                    Require.True(null != hold, "--agency-number requires an argument\n");
                    Require.False(hold.Equals(""), "--agency-number can't use an empty string\n");
                    try
                    {
                        AgencyID = Convert.ToInt32(hold);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex);
                        Console.WriteLine("Agency number (--agency-number) must be a number");
                        Environment.Exit(1);
                    }
                    i += 2;
                }
                else if (Match.CaseInsensitive("--agency-name", args[i]) || Match.CaseInsensitive("-n", args[i]))
                {
                    Agency = args[i + 1];
                    Require.True(null != Agency, "--agency-name requires an argument\n");
                    Require.False(Agency.Equals(""), "--agency-name can't use an empty string\n");
                    i += 2;
                }
                else if (Match.CaseInsensitive("--database-name", args[i]) || Match.CaseInsensitive("-d", args[i]))
                {
                    Database = args[i + 1];
                    Require.True(null != Database, "--database-name requires an argument\n");
                    Require.False(Database.Equals(""), "--database-name can't use an empty string\n");
                    i += 2;
                }
                else if (Match.CaseInsensitive("--audit-mode", args[i]) || Match.CaseInsensitive("-A", args[i]))
                {
                    fresh = false;
                    i += 1;
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

            if(0 == Audit)
            {
                Console.WriteLine("Audit number (--audit ####) must be provided");
                Environment.Exit(1);
            }

            if((0 == AgencyID) && Match.CaseInsensitive(Agency, string.Empty))
            {
                Console.WriteLine("Agency Name (--agency-name %%%%) must be provided\nor\nAgency Number (--agency-number ####) must be provided");
                Environment.Exit(1);
            }

            if(Match.CaseInsensitive(Database, string.Empty))
            {
                Console.WriteLine("Database name (--database-name %%%%) must be provided");
                Environment.Exit(1);
            }

            if(verbose)
            {
                Console.WriteLine(string.Format("starting generation of: {0}", filename));
            }

            Audit_Export gen = new Audit_Export();
            if (0 == AgencyID) AgencyID = gen.Initialize(Database, Agency, fresh);
            else gen.Initialize(Database, fresh);
            gen.Generate(output, AgencyID, Audit, Audit_Name);

            if (verbose)
            {
                Console.WriteLine(string.Format("generation of: {0} is complete", filename));
            }
        }
    }
}
