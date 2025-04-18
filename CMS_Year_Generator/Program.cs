using common;

namespace CMS_Year_Generator
{
    class Program
    {
        static void Main(string[] args)
        {
            int Year = 0;
            List<int> control_years= new();
            string Database = String.Empty;
            string Agency = "ALL";
            List<string> presplit_name = new List<string>();
            bool upgrade = false;
            bool all = false;
            bool verbose = false;
            
            int i = 0;
            while (i < args.Length)
            {
                if (Match.CaseInsensitive("--database-name", args[i]) || Match.CaseInsensitive("-d", args[i]))
                {
                    Database = args[i + 1];
                    Require.True(null != Database, "--database-name requires an argument\n");
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                    Require.False(Database.Equals(""), "--database-name can't use an empty string\n");
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                    i += 2;
                }
                else if (Match.CaseInsensitive("--audit-year", args[i]) || Match.CaseInsensitive("-a", args[i]))
                {
                    string hold = args[i + 1];
                    Require.True(null != hold, "--audit requires an argument\n");
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                    Require.False(hold.Equals(""), "--audit can't use an empty string\n");
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                    try
                    {
                        Year = Convert.ToInt32(hold);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex);
                        Console.WriteLine("Audit number (--audit) must be a number");
                        Environment.Exit(1);
                    }
                    i += 2;
                }
                else if (Match.CaseInsensitive("--audit-controls", args[i]) || Match.CaseInsensitive("-c", args[i]))
                {
                    string raw = args[i + 1];
                    Require.True(null != raw, "--audit-controls requires an argument\n");
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                    Require.False(raw.Equals(""), "--audit-controls can't use an empty string\n");
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                    try
                    {
                        foreach (string hold in raw.Split(','))
                        {
                            control_years.Add(Convert.ToInt32(hold));
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex);
                        Console.WriteLine("Audit number (--audit-controls) must be a number");
                        Environment.Exit(1);
                    }
                    i += 2;
                }
                else if (Match.CaseInsensitive("--update-agency", args[i]))
                {
                    Agency = args[i + 1];
                    Require.True(null != Agency, "--update-agency requires an argument\n");
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                    Require.False(Agency.Equals(""), "--update-agency can't use an empty string\n");
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                    i += 2;
                }
                else if (Match.CaseInsensitive("--inherit-answers-from", args[i]))
                {
                    string hold = args[i + 1];
                    Require.True(null != hold, "--inherit-answers-from requires an argument\n");
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                    Require.False(hold.Equals(""), "--inherit-answers-from can't use an empty string\n");
                    presplit_name.Add(hold);
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                    i += 2;
                }
                else if (Match.CaseInsensitive("--upgrade", args[i]))
                {
                    upgrade = true;
                    i += 1;
                }
                else if (Match.CaseInsensitive("--all", args[i]))
                {
                    all = true;
                    i += 1;
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

            if((0 < presplit_name.Count) && Match.CaseInsensitive("ALL", Agency))
            {
                Console.WriteLine("--inherit-answers-from only works with a single agency, so use --update-agency to indicate which agency you would like to update");
                Environment.Exit(1);
            }

            if(0 == Year)
            {
                Console.WriteLine("Audit number (--audit ####) must be provided");
                Environment.Exit(1);
            }

            if (Match.CaseInsensitive(Database, string.Empty))
            {
                Console.WriteLine("Database name (--database-name %%%%) must be provided");
                Environment.Exit(1);
            }

            if (verbose)
            {
                Console.WriteLine("Beginning Generation");
            }
            // Connect to Database
            CMS_Year_Generator gen = new();
            // Deal with the case of an agency splitting in half
            if(0 < presplit_name.Count) gen.Connect(Database, upgrade, Agency, presplit_name);
            else gen.Connect(Database, upgrade, Agency);

            // build up the list of what controls are being assigned this year
            gen.Get_Agency_Controls(Year, control_years, all);

            // Assign those controls
            gen.Generate(Year);

            // populate the details the report exporters need for this year.
            // if it doesn't already exist
            gen.Populate_report_details(Year);

            if (verbose)
            {
                Console.WriteLine("Generation complete");
            }
        }
    }
}
