using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Data;
using System.Data.SqlClient;
using System.IO;

using System.Configuration;

namespace SQLServerToCSV
{
    class Program
    {
        private static void SQLToCSV(string connectionString, string queryString, 
            string filePath, bool appendOption)
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                using (SqlDataAdapter adapter = new SqlDataAdapter(
                    queryString, connection))
                {
                    DataTable table = new DataTable();
                    adapter.Fill(table);
                    DataTableReader reader = table.CreateDataReader();
                    try
                    {
                        StreamWriter writer = new StreamWriter(filePath, appendOption);

                        for (int i = 0; i < reader.FieldCount - 1; i++)
                            writer.Write(String.Format("{0},", reader.GetName(i)));
                        writer.Write(reader.GetName(reader.FieldCount - 1));
                        writer.Write(writer.NewLine);

                        string colVal = string.Empty;
                        while (reader.Read())
                        {
                            for (int i = 0; i < reader.FieldCount - 1; i++)
                            {
                                colVal = String.Format("{0}", reader.GetValue(i));
                                colVal = StringToCSVCell(colVal);
                                writer.Write(String.Format("{0},", colVal));
                            }
                            writer.Write(reader.GetValue(reader.FieldCount - 1));
                            writer.Write(writer.NewLine);
                        }

                        writer.Close();
                        Console.WriteLine("Execution Successful!");
                        Console.WriteLine(String.Format("The file '{0}' was successfully", filePath));
                        Console.WriteLine("created or overwritten.");
                        Console.WriteLine();
                        Console.WriteLine("Press Enter to Exit.");
                        Console.ReadLine();
                    }
                    catch (IOException)
                    {
                        Console.WriteLine("IOException:");
                        Console.WriteLine("The process cannot access the file");
                        Console.WriteLine(String.Format("'{0}' ", filePath));
                        Console.WriteLine("because it is being used by another process.");
                        Console.WriteLine();
                        Console.WriteLine("Please close the file before running this application.");
                        Console.WriteLine();
                        Console.WriteLine("Press Enter to Exit.");
                        Console.ReadLine();
                    }
                    catch (UnauthorizedAccessException)
                    {
                        Console.WriteLine("UnauthorizedAccessException:");
                        Console.WriteLine(String.Format("Access to the path '{0}' is denied.", filePath));
                        Console.WriteLine();
                        Console.WriteLine("Please check that the following are true:");
                        Console.WriteLine("1. File is not ReadOnly, and");
                        Console.WriteLine("2. You have sufficient privileges to access this file.");
                        Console.WriteLine();
                        Console.WriteLine("You may configure the application to write to another directory or file.");
                        Console.WriteLine();
                        Console.WriteLine("Press Enter to Exit.");
                        Console.ReadLine();
                    }
                }
            }
        }

        public static string StringToCSVCell(string value)
        {
            bool quote = (value.Contains(",") || value.Contains("\"") || 
                value.Contains("\r") || value.Contains("\n"));
            if (quote)
            {
                StringBuilder builder = new StringBuilder();
                builder.Append("\"");
                foreach (char nextChar in value)
                {
                    builder.Append(nextChar);
                    if (nextChar == '"')
                        builder.Append("\"");
                }
                builder.Append("\"");
                return builder.ToString();
            }
            return value;
        }

        static void Main(string[] args)
        {
            AppSettingsReader reader = new AppSettingsReader();

            string dbServer = (string)reader.GetValue("dbServer", typeof(string));
            string dbUser = (string)reader.GetValue("dbUser", typeof(string));
            string dbPwd = (string)reader.GetValue("dbPwd", typeof(string));
            string connectionString = String.Format(
                (string)reader.GetValue("connectionStringFormatString", typeof(string)), 
                dbServer, , dbUser, dbPwd);
            string tableName = (string)reader.GetValue("tableName", typeof(string));
            string queryString = String.Format((string)reader.GetValue("queryStringFormatString", typeof(string)), tableName);
            string winUser = System.Security.Principal.WindowsIdentity.GetCurrent().Name;
            string subPath = (string)reader.GetValue("subPath", typeof(string));
            string fullPath = String.Format(
                (string)reader.GetValue("filePathFormatString", typeof(string)), 
                winUser, subPath, tableName);
            string appendOptionString = (string)reader.GetValue("appendOption", typeof(string));
            bool appendOption = true;
            bool.TryParse(appendOptionString, out appendOption);

            SQLToCSV(connectionString, queryString, fullPath, appendOption);
        }
    }
}
