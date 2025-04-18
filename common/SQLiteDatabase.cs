using System.Data;
using System.Data.SQLite;

namespace common
{
#pragma warning disable IDE0079 // Remove unnecessary suppression
#pragma warning disable CA1822 // Mark members as static
#pragma warning restore IDE0079 // Remove unnecessary suppression

    public class SQLiteDatabase
    {
#pragma warning disable IDE0044 // Add readonly modifier
        private string? dbConnection;
        private SQLiteConnection? cnn;
#pragma warning restore IDE0044 // Add readonly modifier

        public SQLiteDatabase(String inputFile)
        {
            dbConnection = String.Format("Data Source={0}", inputFile);
            cnn = new SQLiteConnection(dbConnection);
            // Make things a bit faster
            cnn.Open();
        }

        public void SQLiteDatabase_Close()
        {
            if (null == cnn) return;
            // Clean up after
            cnn.Close();
        }

        public DataTable GetDataTable(string sql)
        {
            DataTable dt = new();
            try
            {
                SQLiteCommand mycommand = new(cnn)
                {
                    CommandText = sql
                };
                SQLiteDataReader reader = mycommand.ExecuteReader();
                dt.Load(reader);
                reader.Close();
            }
            catch (Exception e)
            {
                throw new Exception(e.Message);
            }
            return dt;
        }

        public int ExecuteNonQuery(string sql)
        {
            SQLiteCommand mycommand = new(cnn)
            {
                CommandText = sql
            };
            try
            {
                int rowsUpdated = mycommand.ExecuteNonQuery();
                return rowsUpdated;
            }
            catch (Exception e)
            {
                throw new Exception(e.Message);
            }
        }

        public string ExecuteScalar(string sql)
        {
            SQLiteCommand mycommand = new(cnn)
            {
                CommandText = sql
            };
            try
            {
                object value = mycommand.ExecuteScalar();
                if (value != null)
                {
#pragma warning disable CS8603 // Possible null reference return.
                    return value.ToString();
#pragma warning restore CS8603 // Possible null reference return.
                }
                return "";
            }
            catch (Exception e)
            {
                throw new Exception(e.Message);
            }
        }

        public int ExecuteIntScalar(string sql)
        {
            SQLiteCommand mycommand = new(cnn)
            {
                CommandText = sql
            };
            try
            {
                object value = mycommand.ExecuteScalar();
                if (value != null)
                {
#pragma warning disable CS8604 // Possible null reference argument.
                    return int.Parse(value.ToString());
#pragma warning restore CS8604 // Possible null reference argument.
                }
                return 0;
            }
            catch (Exception e)
            {
                throw new Exception(e.Message);
            }
        }

        public bool Update(String tableName, Dictionary<String, String> data, String where)
        {
            String vals = "";
            Boolean returnCode = true;
            if (data.Count >= 1)
            {
                foreach (KeyValuePair<String, String> val in data)
                {
                    vals += String.Format(" {0} = '{1}',", val.Key.ToString(), val.Value.ToString());
                }
                vals = vals[..^1];
            }
            try
            {
                ExecuteNonQuery(String.Format("update {0} set {1} where {2};", tableName, vals, where));
            }
            catch
            {
                returnCode = false;
            }
            return returnCode;
        }

        public bool Delete(String tableName, String where)
        {
            Boolean returnCode = true;
            try
            {
                ExecuteNonQuery(String.Format("delete from {0} where {1};", tableName, where));
            }
            catch (Exception fail)
            {
                Console.WriteLine(fail.Message);
                returnCode = false;
            }
            return returnCode;
        }

        public bool Insert(String tableName, Dictionary<String, String> data)
        {
            String columns = "";
            String values = "";
            Boolean returnCode = true;
            foreach (KeyValuePair<String, String> val in data)
            {
                // but don't allow ' in column names
                columns += String.Format(" '{0}',", val.Key.ToString());
                // Work around ' in fields
                values += String.Format(" '{0}',", val.Value.Replace("'", "''"));
            }
            columns = columns[..^1];
            values = values[..^1];
            try
            {
                string expression = String.Format("insert into {0}({1}) values({2});", tableName, columns, values);
                ExecuteNonQuery(expression);
            }
            catch (Exception fail)
            {
                Console.WriteLine(fail.Message);
                returnCode = false;
            }
            return returnCode;
        }
    }
#pragma warning disable IDE0079 // Remove unnecessary suppression
#pragma warning restore CA1822 // Mark members as static
#pragma warning restore IDE0079 // Remove unnecessary suppression
}
