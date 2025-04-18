using common;
using OfficeOpenXml;

namespace CMS_Audit_Import
{
    internal class Legacy_Import
    {
        public SQLiteDatabase? mydatabase;
        private ExcelWorkbook? book;
        private ExcelWorksheet? currentWorksheet;

        public void Initialize(string Database, FileInfo existingFile)
        {
            book = new ExcelPackage(existingFile).Workbook;

            if (!File.Exists(Database))
            {
                Console.WriteLine("Unable to find " + Database);
                return;
            }

            // Get the first worksheet
            currentWorksheet = book.Worksheets.First();

            // Connect to Database
            mydatabase = new SQLiteDatabase(Database);
        }
        public void Cleanup()
        {
            Require.False(null == mydatabase, "impossible cleanup state hit");
#pragma warning disable CS8602 // Dereference of a possibly null reference.
            mydatabase.SQLiteDatabase_Close();
#pragma warning restore CS8602 // Dereference of a possibly null reference.
        }

        public void Precursor_Import(versions Version)
        {
            if (null == Version) return;
            if (null == Version.Year) return;
            if (null == Version.Agency) return;
            if (null == book) return;
            int Year = (int)Version.Year;
            string Agency = Version.Agency;
            foreach (ExcelWorksheet sheet in book.Worksheets)
            {
                // Skip First Sheet
                if (null == currentWorksheet) continue;
                if (currentWorksheet.Equals(sheet)) continue;
                if (sheet.Name.Equals("Form")) continue;

                // Todo remove after 2026
                if (Version.ancientp())
                {
                    if (sheet.Name.Equals("Sheet1")) continue;
                    if (sheet.Name.Equals("2015 Evidence")) continue;
                    if (sheet.Name.Equals("Evidence")) continue;
                    if (sheet.Name.Equals("Query")) continue;
                }

                // Check if we have this control yet
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                object temp = mydatabase.ExecuteScalar(string.Format("SELECT CONTROL FROM CONTROLS WHERE CONTROL = \"{0}\";", sheet.Name));
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                if ("" == temp.ToString())
                {
                    Console.WriteLine("Skipping import of control " + sheet.Name);
                    continue;
                }

                string Related_to = Find_field_harder(1, 2, sheet, "Related to App or Agency Service", 1);
                string Describe_implemented = "";
                if (Version.ancientp())
                {
                    Describe_implemented = Find_field_harder(1, 2, sheet, "Describe how the control is implemented or why it is not applicable.", 1);
                }
                else if(Version.pregeneratedp())
                {
                    Describe_implemented = Find_field(1, 2, sheet, "Description of how the control is implement- ed or why not applicable. (Enter/Update)");
                }
                else if (Version.prototypep())
                {
                    object? hold = sheet.Cells["B28"].Value;
                    Require.False(null == hold, "Bad description");
#pragma warning disable CS8602 // Dereference of a possibly null reference.
#pragma warning disable CS8600 // Converting null literal or possible null value to non-nullable type.
                    Describe_implemented = hold.ToString();
#pragma warning restore CS8600 // Converting null literal or possible null value to non-nullable type.
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                }
                else
                {
                    Console.WriteLine("Aborting for safety");
                    Environment.Exit(1);
                }

                string Describe_Evidence = Find_field_harder(1, 2, sheet, "Describe the Evidence that is available", 1);
                string How_Evidence = Find_field_harder(1, 2, sheet, "How will the evidence be provided", 1);
                string Evidence_Source = Find_field_harder(1, 2, sheet, "Source of evidence", 1);

                string Preaudit_status = "";
                string Preaudit_Desc_Date = "" ;
                string Evidence_Submit_Date = "";
                if (Version.ancientp())
                {
                    Preaudit_status = "Gap";
                    Preaudit_Desc_Date = Find_date_field_harder(1, 2, sheet, "Date Pre-Audit Review Completed", 1);
                    Evidence_Submit_Date = "03/24/2016";
                }
                else if (Version.pregeneratedp())
                {
                    Preaudit_status = Find_field(1, 2, sheet, "Status");
                    Preaudit_Desc_Date = Find_date_field(1, 2, sheet, "Date Desc of Implementation Complete");
                    Evidence_Submit_Date = Find_date_field(1, 2, sheet, "Date Evidence Submitted");
                }
                else if (Version.prototypep())
                {
                    object? hold = sheet.Cells["B39"].Value;
#pragma warning disable CS8602 // Dereference of a possibly null reference.
#pragma warning disable CS8600 // Converting null literal or possible null value to non-nullable type.
                    Require.False(null == hold, "Bad Preaudit_Status");
                    Preaudit_status = hold.ToString();

                    hold = sheet.Cells["B33"].Value;
                    Require.False(null == hold, "Bad Preaudit_Desc_Date");
                    Preaudit_Desc_Date = hold.ToString();

                    hold = sheet.Cells["B34"].Value;
                    Require.False(null == hold, "Bad Evidence_Submit_Date");
                    Evidence_Submit_Date = hold.ToString();
#pragma warning restore CS8600 // Converting null literal or possible null value to non-nullable type.
#pragma warning restore CS8602 // Dereference of a possibly null reference.

                }
                else
                {
                    Console.WriteLine("Aborting for safety");
                    Environment.Exit(1);
                }
                string Audit_Status = "";
                if (Version.ancientp())
                {
                    Audit_Status = Find_field_harder(1, 2, sheet, "STATUS", 1);
                }
                else if (Version.pregeneratedp())
                {
                    Audit_Status = Find_field(1, 2, sheet, "STATUS of Pre-Audit Review");
                }
                else
                {
                    Console.WriteLine("Aborting for safety");
                    Environment.Exit(1);
                }

                string Status = Find_field(1, 2, sheet, "Status");
                string Criticality = Find_field(1, 2, sheet, "Criticality");

                // Get Universal ControlID
                string sql = string.Format("SELECT CONTROL_ID FROM CONTROLS WHERE CONTROL = '{0}' AND AUDIT_TYPE = {1};", sheet.Name, 31337);
                string controlID = mydatabase.ExecuteScalar(sql);

                // Finally get the last Audit Date
                string LastAudit;
                if (Version.ancientp())
                {
                    LastAudit = "2015";
                }
                else
                {
                    Preaudit_status = mydatabase.ExecuteScalar(String.Format("SELECT AUDIT_STATUS FROM AUDIT_CONTROLS WHERE CONTROL_ID = \"{0}\" AND AGENCY_ID = \"{1}\" ORDER BY AUTO_ID DESC LIMIT 1;", controlID, Agency));
                    if (Version.pregeneratedp() && ((null == Preaudit_status) || Match.CaseInsensitive(Preaudit_status, "")))
                    {
                        Preaudit_status = Find_field(1, 2, sheet, "Status");
                    }
                    LastAudit = mydatabase.ExecuteScalar(String.Format("SELECT AUDIT FROM AUDIT_CONTROLS WHERE CONTROL_ID = \"{0}\" AND AGENCY_ID = \"{1}\" ORDER BY AUDIT DESC LIMIT 1;", controlID, Agency));
                }

                // CREATE TABLE Audit_Controls (AUTO_ID INTEGER NOT NULL PRIMARY KEY AUTOINCREMENT UNIQUE, AUDIT INTEGER, AGENCY_ID, CONTROL_ID, RELATED_TO, IMPLEMENTATION, EVIDENCE_DESCRIPTION, EVIDENCE_DELIVERY, EVIDENCE_FILES, PREAUDIT_STATUS, PREAUDIT_REVIEW_DATE, EVIDENCE_SUBMIT_DATE, STATUS, CRITICALITY, LAST_AUDIT INTEGER);
#pragma warning disable CS8604 // Possible null reference argument.
                Dictionary<string, string> tmp = new() {{"AUDIT", Year.ToString()},
                                                        {"AGENCY_ID", Agency},
                                                        {"CONTROL_ID", controlID},
                                                        {"RELATED_TO", Related_to},
                                                        {"IMPLEMENTATION", Describe_implemented},
                                                        {"EVIDENCE_DESCRIPTION", Describe_Evidence},
                                                        {"EVIDENCE_DELIVERY", How_Evidence},
                                                        {"EVIDENCE_FILES", Evidence_Source},
                                                        {"PREAUDIT_STATUS", Preaudit_status},
                                                        {"PREAUDIT_REVIEW_DATE", Preaudit_Desc_Date},
                                                        {"EVIDENCE_SUBMIT_DATE", Evidence_Submit_Date},
                                                        {"AUDIT_STATUS", Audit_Status},
                                                        {"STATUS", Status},
                                                        {"CRITICALITY", Criticality},
                                                        {"LAST_AUDIT", LastAudit}
                };
#pragma warning restore CS8604 // Possible null reference argument.
                mydatabase.Insert("AUDIT_CONTROLS", tmp);
            }
        }

        public static string Find_field(int col_match, int col_return, ExcelWorksheet currentWorksheet, string field_text)
        {
            string? ret = null;
            try
            {
                for (int i = 1; i <= 100; i += 1)
                {
                    if (null != ret) break;

                    string field = (string)currentWorksheet.Cells[i, col_match].Value;
                    if (null == field)
                    {
                        continue;
                    }
                    else if ((null == ret) && field.Trim().Equals(field_text))
                    {
                        ret = (string)currentWorksheet.Cells[i, col_return].Value;
                        if (null == ret) ret = "";
                    }
                }
            }
            catch
            {
                Console.WriteLine(string.Format("Unable to find field {0} in sheet {1}", field_text, currentWorksheet.Name));
                Environment.Exit(1);
            }

            if (null == ret) return "";
            return ret.Replace('"', '\'');
        }

        public static string Find_field_harder(int col_match, int col_return, ExcelWorksheet currentWorksheet, string field_text, int skip)
        {
            string? ret = null;
            try
            {
                for (int i = 1; i <= 100; i += 1)
                {
                    if (null != ret) break;

                    string field = (string)currentWorksheet.Cells[i, col_match].Value;
                    if (null == field)
                    {
                        continue;
                    }
                    else if ((null == ret) && field.Trim().Equals(field_text))
                    {
                        ret = (string)currentWorksheet.Cells[i, col_return].Value;
                        if (0 != skip)
                        {
                            skip -= 1;
                            ret = null;
                        }
                        else if (null == ret) ret = "";
                    }
                }
            }
            catch
            {
                Console.WriteLine(string.Format("Unable to find field {0} in sheet {1}", field_text, currentWorksheet.Name));
                Environment.Exit(1);
            }

            if (null == ret) return "";
            return ret.Replace('"', '\'');
        }

        public static string Find_date_field(int col_match, int col_return, ExcelWorksheet currentWorksheet, string field_text)
        {
            string? ret = null;
            try
            {
                for (int i = 1; i <= 100; i += 1)
                {
                    if (null != ret) break;

                    string field = (string)currentWorksheet.Cells[i, col_match].Value;
                    if (null == field)
                    {
                        continue;
                    }
                    else if ((null == ret) && field.Equals(field_text))
                    {
                        if (null == currentWorksheet.Cells[i, col_return].Value)
                        {
                            return "";
                        }

#pragma warning disable CS8602 // Dereference of a possibly null reference.
                        string[] tokens = currentWorksheet.Cells[i, col_return].Value.ToString().Split(' ');
#pragma warning restore CS8602 // Dereference of a possibly null reference.

                        ret = tokens[0];
                        if (!ret.Contains('/'))
                        {
                            ret = currentWorksheet.Cells[i, col_return].Text;
                        }
                    }
                }
            }
            catch
            {
                Console.WriteLine(string.Format("Unable to find field {0} in sheet {1}", field_text, currentWorksheet.Name));
                Environment.Exit(1);
            }

            if (null == ret) return "Unable to find date";
            return ret;
        }

        public static string Find_date_field_harder(int col_match, int col_return, ExcelWorksheet currentWorksheet, string field_text, int skip)
        {
            string? ret = null;
            object? hold;
            try
            {
                for (int i = 1; i <= 100; i += 1)
                {
                    if (null != ret) break;

                    string field = (string)currentWorksheet.Cells[i, col_match].Value;
                    if (null == field)
                    {
                        continue;
                    }
                    else if ((null == ret) && field.Equals(field_text))
                    {
                        hold = currentWorksheet.Cells[i, col_return].Value;
                        if (0 != skip)
                        {
                            skip -= - 1;
                            ret = null;
                        }
                        else if (null == hold) ret = "";
                        else
                        {
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                            string[] tokens = currentWorksheet.Cells[i, col_return].Value.ToString().Split(' ');
#pragma warning restore CS8602 // Dereference of a possibly null reference.

                            ret = tokens[0];
                            if (!ret.Contains('/'))
                            {
                                ret = currentWorksheet.Cells[i, col_return].Text;
                            }
                        }
                    }
                }
            }
            catch
            {
                Console.WriteLine(string.Format("Unable to find field {0} in sheet {1}", field_text, currentWorksheet.Name));
                Environment.Exit(1);
            }

            if (null == ret) return "Unable to find date";
            return ret;
        }
    }
}
