using common;
using OfficeOpenXml;

namespace CMS_Audit_Import
{
    internal class Audit_Import
    {
        public SQLiteDatabase? mydatabase;
        private ExcelWorkbook? book;
        private ExcelWorksheet? currentWorksheet;

        public versions Initialize(string Database, FileInfo existingFile, bool legacymode)
        {
            book = new ExcelPackage(existingFile).Workbook;
            versions? r = null;

            try
            {
                // Get the first worksheet
                currentWorksheet = book.Worksheets.First();
                // Read Agency and Year
                r = Read_FirstPage(currentWorksheet, legacymode);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                Console.WriteLine("Unable to read Excel contents");
                Environment.Exit(1);
            }

            // Connect to Database
            mydatabase = new SQLiteDatabase(Database);
            if (null != r) return r;
            Console.WriteLine("unable to figure out version");
            Environment.Exit(1);
            return null;
        }

        public void Cleanup()
        {
            Require.False(null == mydatabase, "impossible cleanup state hit");
#pragma warning disable CS8602 // Dereference of a possibly null reference.
            mydatabase.SQLiteDatabase_Close();
#pragma warning restore CS8602 // Dereference of a possibly null reference.
        }

        public versions Read_FirstPage(ExcelWorksheet currentWorksheet, bool legacymode)
        {
            string header = Read_cell("A1", currentWorksheet);
            string[] hold1 = header.Split(' ');
            string year = hold1[0];
            int Year = int.Parse(year);
            string Agency;
            if (legacymode)
            {
                Agency = Read_number("K1", currentWorksheet).ToString();
                return new versions(31337, Agency, Year);
            }
            else
            {
                Agency = Read_number("F4", currentWorksheet).ToString();
                int Audit_Version = Read_number("F5", currentWorksheet);
                Require.False(Match.CaseInsensitive(Audit_Version.ToString(), string.Empty), "please use --legacy-import");
                Require.False(Match.CaseInsensitive(Audit_Version.ToString(), "0"), "please use --legacy-import");
                return new versions(Audit_Version, Agency, Year);
            }
        }

        public void Import(versions v)
        {
            if (v.arc_ampep()) ARC_AMPE_Audit_Import.ARC_AMPE_Import(mydatabase, v.Agency, book, v.Year, v.arc_ampe);
            else if (v.cms_22p()) CMS_Audit_Import.CMS_22_Import(mydatabase, v.Agency, book, v.Year, v.cms_22);
            else if (v.prototypep()) CMS_Audit_Import.CMS_22_Import(mydatabase, v.Agency, book, v.Year, v.prototype);
            else
            {
                Console.WriteLine("version: {0} not supported", v.debug());
                Environment.Exit(1);
            }
        }

        private static string Read_cell(string address, ExcelWorksheet currentWorksheet)
        {
            return currentWorksheet.Cells[address].Text.Replace('"', '\'');
        }

        private static int Read_number(string address, ExcelWorksheet currentWorksheet)
        {
            try
            {
                object result = currentWorksheet.Cells[address].Value;
                double hold = Convert.ToDouble(result);
                return (int)Math.Ceiling(hold);
            }
            catch
            {
                Console.WriteLine("Field " + address + " does not contain a valid number");
                Environment.Exit(3);
            }
            return -1;
        }
    }
}
