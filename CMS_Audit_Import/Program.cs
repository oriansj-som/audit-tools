using common;

namespace CMS_Audit_Import
{
    class Program
    {
        static void Main(string[] args)
        {

            // Get the file we are going to process
            FileInfo? existingFile = null;
            // Target File we are updating
            string Database = String.Empty;
            bool verbose = false;
            // Figure out how to import
            versions version = new();

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
                else if (Match.CaseInsensitive("--legacy-import", args[i]))
                {
                    string generation = args[i + 1];
                    Require.True(null != generation, "--legacy-import requires a generation");
                    try
                    {
                        int geny = int.Parse(generation);
                        version = new versions(geny);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.ToString());
                        Environment.Exit(1);
                    }
                    i += 2;
                }
                else if (Match.CaseInsensitive("--special-import", args[i]))
                {
                    string generation = args[i + 1];
                    Require.True(null != generation, "--special-import requires a generation");
                    string agent = args[i + 2];
                    Require.True(null != agent, "--special-import requires an Agency ID code");
                    string year = args[i + 3];
                    Require.True(null != year, "--special-import requires a year");
                    try
                    {
#pragma warning disable CS8604 // Possible null reference argument.
                        int genny = Int32.Parse(generation);
                        int Year = Int32.Parse(year);
                        version = new versions(genny,agent, Year);
#pragma warning restore CS8604 // Possible null reference argument.
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                        Environment.Exit(1);
                    }
                    
                    i += 4;
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

            if(null == existingFile)
            {
                Console.WriteLine("input file (--input-file %%%%) is required");
                Environment.Exit(1);
            }

            if(Match.CaseInsensitive(Database, string.Empty))
            {
                Console.WriteLine("Database file (--database-name %%%%) is required");
                Environment.Exit(1);
            }

            if(verbose)
            {
                Console.WriteLine("Beginning import");
            }

            if (version.arc_ampep() || version.cms_22p() || version.prototypep())
            {
                Audit_Import gen = new();
                version = gen.Initialize(Database, existingFile, version.prototypep());
                gen.Import(version);
                gen.Cleanup();
            }
            else
            {
                Legacy_Import gen = new();
                gen.Initialize(Database, existingFile);
                gen.Precursor_Import(version);
                gen.Cleanup();

            }

            if (verbose)
            {
                Console.WriteLine("Import Complete");
            }
        }
    }
}
