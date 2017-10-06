using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;

namespace ConsoleApp1
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hallo, es wird bald etwas ausgegeben!");
            // Provider = Microsoft.Jet.OLEDB.4.0; Data Source = "C:\Users\Reg1nleifr\Desktop\Koch Ines\DBKesztele.mdb"
            var connectionString = "Provider = Microsoft.Jet.OLEDB.4.0; Data Source = \"C:\\Users\\Reg1nleifr\\Desktop\\Koch Ines\\DBKesztele.mdb\"";
            var queryString = "select * from Debitor where Month(Birthdate) = " + DateTime.Now.Month;

            using (OleDbConnection connection = new OleDbConnection(connectionString))
            using (OleDbCommand command = new OleDbCommand(queryString, connection))
            {
                try
                {
                    connection.Open();
                    OleDbDataReader reader = command.ExecuteReader();
                    while (reader.Read())
                    {
                        Console.WriteLine(reader["Firstname"].ToString());
                    }
                    reader.Close();
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }

            // Aka - ur komplexer join für alle Daten nach 6 Monaten
            queryString = @"SELECT * FROM 
                           (RechnungArtikel ra 
                            INNER JOIN Rechnung r ON r.ID = ra.RechnungsID)
                            INNER JOIN Debitor d ON d.ID = int(r.DebitorID)
                                WHERE ProductID like '00000125'
                                AND DateDiff('m', Time, Date()) >= 6";
            using (OleDbConnection connection = new OleDbConnection(connectionString))
            using (OleDbCommand command = new OleDbCommand(queryString, connection))
            {
                try
                {
                    connection.Open();
                    OleDbDataReader reader = command.ExecuteReader();
                    while (reader.Read())
                    {
                        Console.WriteLine("Bezeichnung: {0} at {1}, by {2}", reader["Bezeichnung"].ToString(), reader["Time"], reader["Firstname"]);

                    }
                    reader.Close();
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }

            Console.ReadKey();
        }
    }
}
