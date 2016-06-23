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
        static string connectionString = String.Format(
            "data source={0};initial catalog={1};User ID={2};Password={3}", 
            dbServer, XXXX, dbUser, dbPwd);
        static string queryString = "SELECT XXXX FROM XXXX WHERE XXXX AND XXXX";
        static string winUser = System.Security.Principal.WindowsIdentity.GetCurrent().Name;
        static string filePath = String.Format("C:\\Users\\{0}\\XXXX\\XXXX", winUser);

        private static void SQLToCSV()
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                SqlCommand command = new SqlCommand(queryString, connection);
                command.Connection.Open();
                SqlDataReader reader = command.ExecuteReader();
                try
                {
                    StreamWriter writer = new StreamWriter(filePath, false);
                    for (int i = 0; i < reader.FieldCount - 1; i++)
                        writer.Write(String.Format("{0},", reader.GetName(i)));
                    writer.Write(reader.GetName(reader.FieldCount - 1));
                    writer.Write(writer.NewLine);

                    while (reader.Read())
                    {
                        for (int i = 0; i < reader.FieldCount - 1; i++)
                            writer.Write(String.Format("{0},", reader.GetValue(i)));
                        writer.Write(reader.GetValue(reader.FieldCount - 1));
                        writer.Write(writer.NewLine);
                    }

                    writer.Close();
                }
                catch (IOException)
                {
                    Console.WriteLine("IOException: Please close the file!");
                    Console.ReadLine();
                }
            }
        }

        static void Main(string[] args)
        {
            SQLToCSV();
        }
    }
}
