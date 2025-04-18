using common;

namespace CMS_Subset_Update
{
    class Program
    {
        static readonly int version = 31340;
        static void Main(string[] args)
        {
            // What Audit Number to use
            int Audit = 0;
            // What Agency
            string Agency = string.Empty;
            int AgencyID = 0;
            // What controls to export
            List<string> set = new();
            // Target File we are updating
            string Database = string.Empty;
            bool verbose = false;
            bool upgrade = false;

            int i = 0;
            while (i < args.Length)
            {
                if (Match.CaseInsensitive("--audit", args[i]) || Match.CaseInsensitive("-a", args[i]))
                {
                    string hold = args[i + 1];
                    Require.True(null != hold, "--audit requires an argument\n");
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                    Require.False(hold.Equals(""), "--audit can't use an empty string\n");
#pragma warning restore CS8602 // Dereference of a possibly null reference.
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
                else if (Match.CaseInsensitive("--agency-number", args[i]) || Match.CaseInsensitive("-N", args[i]))
                {
                    string hold = args[i + 1];
                    Require.True(null != hold, "--agency-number requires an argument\n");
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                    Require.False(hold.Equals(""), "--agency-number can't use an empty string\n");
#pragma warning restore CS8602 // Dereference of a possibly null reference.
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
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                    Require.False(Agency.Equals(""), "--agency-name can't use an empty string\n");
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
                else if (Match.CaseInsensitive("--controls", args[i]))
                {
                    string hold = args[i + 1];
                    Require.True(null != hold, "--controls requires an argument\n");
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                    Require.False(hold.Equals(""), "--controls can't use an empty string\n");
#pragma warning restore CS8602 // Dereference of a possibly null reference.

                    set = hold.Split(',').ToList();
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

            if (0 == set.Count)
            {
                Console.WriteLine("Controls (--controls '%%%','%%%') must be provided");
                Environment.Exit(1);
            }

            if (0 == Audit)
            {
                Console.WriteLine("Audit number (--audit ####) must be provided");
                Environment.Exit(1);
            }

            if ((0 == AgencyID) && Match.CaseInsensitive(Agency, string.Empty))
            {
                Console.WriteLine("Agency Name (--agency-name %%%%) must be provided\nor\nAgency Number (--agency-number ####) must be provided");
                Environment.Exit(1);
            }

            if (Match.CaseInsensitive(Database, string.Empty))
            {
                Console.WriteLine("Database name (--database-name %%%%) must be provided");
                Environment.Exit(1);
            }

            if (verbose)
            {
                Console.WriteLine("starting Update of database");
            }

            CMS_Subset_Generator gen = new();
            if (0 == AgencyID) AgencyID = gen.Initialize(Database, Agency, upgrade);
            else gen.Initialize(Database, upgrade);
            gen.Get_Agency_Controls(AgencyID, set, version);
            gen.Generate(Audit);

            if (verbose)
            {
                Console.WriteLine("Database update is complete");
            }

        }
    }
}
