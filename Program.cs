using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Data.SqlClient;
using System.IO;

namespace SQLServerToCSV
{
    class Program
    {
        static string dbServer = "XXXX";
        static string dbUser = "XXXX";
        static string dbPwd = "XXXX";
        static string connectionTemplate = String.Format(
            "data source={0};initial catalog={1};User ID={2};Password={3}", 
            dbServer, XXXX, dbUser, dbPwd);
        static string sqlQuery = "SELECT XXXX FROM XXXX WHERE XXXX AND XXXX";
        static string winUser = System.Security.Principal.WindowsIdentity.GetCurrent().Name;
        static string filePath = String.Format("C:\\Users\\{0}\\XXXX", winUser);

        private static string SQLToString(string queryString, string connectionString)
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                SqlCommand command = new SqlCommand(queryString, connection);
                command.Connection.Open();
                SqlDataReader reader = command.ExecuteReader();
                StringBuilder builder = new StringBuilder();
                try
                {
                    builder.Append(String.Format("{0}", reader.GetName(0)));
                    for (int i = 1; i < reader.FieldCount; i++)
                        builder.Append(String.Format(",{0}", reader.GetName(i)));
                    while (reader.Read())
                    {
                        for (int i = 0; i < reader.FieldCount; i++)
                            builder.Append(String.Format(",{0}", reader.GetValue(i)));
                    }
                }
                finally
                {
                    reader.Close();
                }
                return builder.ToString();
            }
        }

        private static void StringToFile(string csv, string path)
        {
            StreamWriter writer = new StreamWriter(path, false);
            writer.Write(csv);
        }

        static void Main(string[] args)
        {
            string dbCSV = SQLToString(sqlQuery, connectionTemplate);
            Console.WriteLine(dbCSV);
            StringToFile(dbCSV, filePath);
        }
    }
}
