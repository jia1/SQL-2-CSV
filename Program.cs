using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Data;
using System.Data.SqlClient;
using System.IO;

namespace SQLServerToCSV
{
    class Program
    {
        private static string dbServer = "";
        private static string dbUser = "";
        private static string dbPwd = "";
        private static string connectionString = String.Format(
            "data source={0};initial catalog={1};User ID={2};Password={3}", 
            dbServer, , dbUser, dbPwd);
        private static string tableName = "";
        private static string queryString = String.Format("SELECT * FROM {0}", tableName);
        private static string winUser = System.Security.Principal.WindowsIdentity.GetCurrent().Name;
        private static string filePath = String.Format("C:\\Users\\{0}\\Desktop\\{1}.csv", winUser, tableName);
        private static bool appendOption = false;
        
        private static void SQLToCSV()
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
            SQLToCSV();
        }
    }
}
