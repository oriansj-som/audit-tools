using common;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace Excel_Importer
{
    internal class column
    {
        public string name;
        public bool pretty;
    }
    internal class Excel_Tab_Importer
    {
        private ExcelWorksheet ws;
        private bool verbose;
        private List<column> columns;

        public string initialize(string name, ExcelWorksheet w, bool v)
        {
            verbose = v;
            ws = w;
            return create_table(name);
        }

        private string create_table(string name)
        {
            columns= new List<column>();
            string fields = string.Empty;
            int col = 1;
            int row = 1;
            while(true)
            {
                string hold = ws.Cells[row, col].Text;
                if (hold.Equals(string.Empty)) break;
                string[] split = hold.Split(new Char[] { ' ', '\n', '[', ']' }, StringSplitOptions.RemoveEmptyEntries);

                if (null == split)
                {
                    if (verbose)
                    {
                        Console.WriteLine(string.Format("row: {0}, col: {1} bad", row, col));
                    }
                    Environment.Exit(1);
                }
                else if (0 == split.Length)
                {
                    if (verbose)
                    {
                        Console.WriteLine(string.Format("row: {0}, col: {1} empty", row, col));
                    }
                    break;
                }

                if(1 < col)
                {
                    fields += " ,";
                }

                fields = fields + "`" + split[0].ToUpper() + "` ";
                column foo = new column();
                foo.pretty = false;
                foo.name = split[0];

                int i = 1;
                while (i < split.Length)
                {
                    if (!Match.CaseInsensitive("PRETTY", split[i])) fields = fields + " " + split[i].ToUpper();
                    else foo.pretty = true;

                    i += 1;
                }
                columns.Add(foo);
                col += 1;

            }

            return "CREATE TABLE IF NOT EXISTS `" + name + "` ( " + fields + " );";
        }

        public List<Dictionary<String, String>> import()
        {
            List<Dictionary<String, String>> values = new List<Dictionary<String, String>>();
            int row = 2;
            while (true)
            {
                int col = 1;
                Dictionary<String, String> value = new Dictionary<String, String>();

                foreach (column n in columns)
                {
                    string name = n.name;
                    string hold;
                    if (n.pretty) hold = ugly.make_ugly(ws.Cells[row, col].RichText);
                    else hold = ws.Cells[row, col].Text;
                    if ((1 == col) && hold.Equals(string.Empty)) goto escape;

                    if (!value.ContainsKey(name))
                    {
                        value.Add(name, hold);
                    }
                    else
                    {
                        Console.WriteLine(string.Format("Duplicate Column Name {0} found, duplicate column names are not allowed", name));
                        Environment.Exit(3);
                    }
                    col += 1;
                }
                if (verbose) 
                {
                    Console.WriteLine(String.Format("read up to line {0}", row));
                }

                row += 1;
                values.Add(value);
            }

escape:
            return values;
        }
    }
}
