using System.Collections.Generic;
using System.Data;
using System.Data.Odbc;
using System.Dynamic;
using CsvHelper;
using System.IO;
using System.Globalization;

namespace QueryToCSV
{
    /// <summary>
    /// Helper class for exporting ODBC query results to a csv file.
    /// </summary>
    public static class QueryCSVHelper
    {
        private static void AddProperty(ExpandoObject expando, string propertyName, object propertyValue)
        {
            var expandoDict = expando as IDictionary<string, object>;
            if (expandoDict.ContainsKey(propertyName))
            {
                expandoDict[propertyName] = propertyValue;
            }
            else
            {
                expandoDict.Add(propertyName, propertyValue);
            }
        }

        /// <summary>
        /// Take a connection string, a query string, and a file path, and write the results of the query into a csv file at the location referred to in the file path. 
        /// </summary>
        /// <param name="connectionString">Connection string to connect to your odbc service.</param>
        /// <param name="queryString">Query to be ran on the database connected to with connectionString</param>
        /// <param name="csvOutPath">The filepath where the csv should be generated.</param>
        public static void OutputCsv(string connectionString, string queryString, string csvOutPath)
        {
            var command = new OdbcCommand(queryString);

            using (var connection = new OdbcConnection(connectionString))
            {
                command.Connection = connection;
                connection.Open();
                var data_reader = command.ExecuteReader();
                data_reader.Read();

                var tableSchema = data_reader.GetSchemaTable();
                var field_count = data_reader.FieldCount;
                var enumerated_columns = new Dictionary<int, string> { };

                var i = 0;
                foreach (DataRow row in tableSchema.Rows)
                {
                    enumerated_columns.Add(i, row["ColumnName"].ToString());
                    i++;
                }
                
                var query_rows = new List<dynamic> { };

                while (data_reader.Read())
                {
                    dynamic row = new ExpandoObject();
                    for (var x = 0; x < data_reader.FieldCount; x++)
                    {
                        AddProperty(row, enumerated_columns[x], data_reader[x]);
                    }
                    query_rows.Add(row);
                }

                using (var writer = new StringWriter())
                using (var csv = new CsvWriter(writer, CultureInfo.InvariantCulture))
                {
                    csv.WriteRecords(query_rows);
                    File.WriteAllText(csvOutPath, writer.ToString());

                }

            }
        }

        /// <summary>
        /// Take a connection string, a query string, and a file path, and write the results of the query into a csv file at the location referred to in the file path. 
        /// Writes the file asynchronously.
        /// </summary>
        /// <param name="connectionString">Connection string to connect to your odbc service.</param>
        /// <param name="queryString">Query to be ran on the database connected to with connectionString</param>
        /// <param name="csvOutPath">The filepath where the csv should be generated.</param>
        public static async void OutputCsvAsync(string connectionString, string queryString, string csvOutPath)
        {
            var command = new OdbcCommand(queryString);

            using (var connection = new OdbcConnection(connectionString))
            {
                command.Connection = connection;
                connection.Open();
                var data_reader = command.ExecuteReader();
                data_reader.Read();

                var tableSchema = data_reader.GetSchemaTable();
                var field_count = data_reader.FieldCount;
                var enumerated_columns = new Dictionary<int, string> { };

                var i = 0;
                foreach (DataRow row in tableSchema.Rows)
                {
                    enumerated_columns.Add(i, row["ColumnName"].ToString());
                    i++;
                }

                var query_rows = new List<dynamic> { };

                while (data_reader.Read())
                {
                    dynamic row = new ExpandoObject();
                    for (var x = 0; x < data_reader.FieldCount; x++)
                    {
                        AddProperty(row, enumerated_columns[x], data_reader[x]);
                    }
                    query_rows.Add(row);
                }

                using (var writer = new StringWriter())
                using (var csv = new CsvWriter(writer, CultureInfo.InvariantCulture))
                {
                    csv.WriteRecords(query_rows);
                    await File.WriteAllTextAsync(csvOutPath, writer.ToString());

                }

            }
        }

        /// <summary>
        /// Take a connection string and query string, and return the results of the ODBC query as a list of Exapndo objects.
        /// </summary>
        /// <param name="connectionString">Connection string to connect to your odbc service.</param>
        /// <param name="queryString">Query to be ran on the database connected to with connectionString</param>
        /// <returns>A list of dynamic objects each representing a single row, the values in the row are mapped to properties with the same names as the column names in the database you are querying from.</returns>
        public static List<dynamic> GetData(string connectionString, string queryString)
        {
            var command = new OdbcCommand(queryString);

            using (var connection = new OdbcConnection(connectionString))
            {
                command.Connection = connection;
                connection.Open();
                var data_reader = command.ExecuteReader();
                data_reader.Read();

                var tableSchema = data_reader.GetSchemaTable();
                var field_count = data_reader.FieldCount;
                var enumerated_columns = new Dictionary<int, string> { };

                var i = 0;
                foreach (DataRow row in tableSchema.Rows)
                {
                    enumerated_columns.Add(i, row["ColumnName"].ToString());
                    i++;
                }

                var query_rows = new List<dynamic> { };

                while (data_reader.Read())
                {
                    dynamic row = new ExpandoObject();
                    for (var x = 0; x < data_reader.FieldCount; x++)
                    {
                        AddProperty(row, enumerated_columns[x], data_reader[x]);
                    }
                    query_rows.Add(row);
                }

                return query_rows;

            }
        }
    }
}
