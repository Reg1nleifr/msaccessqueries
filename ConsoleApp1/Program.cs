using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using System.Windows.Forms;
using System.IO;

namespace ConsoleApp1
{
    class Program
    {   
        static string[] months = { "Jänner", "Februar", "März", "April", "Mai", "Juni", "Juli", "August", "September", "Oktober", "November", "Dezember" };
        static Dictionary<string, string> data = new Dictionary<string, string>();

        [STAThread]
        static void Main(string[] args)
        {
            if(ReadDataFromFile("./db.ini"))
            {
                CreateBirthDayCSV();
                CreateBloodCSV();
            }
            Console.ReadKey();
        }

        private static void CreateBirthDayCSV()
        {
            var csv = new StringBuilder();
            csv.AppendLine(data["csvHeaderBirthdate"]);
            var queryString = data["birthdayQuery"] + DateTime.Now.Month;

            using (OleDbConnection connection = new OleDbConnection(data["connectionString"]))
            using (OleDbCommand command = new OleDbCommand(queryString, connection))
            {
                try
                {
                    connection.Open();
                    OleDbDataReader reader = command.ExecuteReader();
                    while (reader.Read())
                    {
                        var birthdate = ((DateTime)reader[data["birthdayColName"]]);
                        var age = DateTime.Now.Year - birthdate.Year;
                        var firstname = reader[data["firstnameColName"]].ToString();
                        var lastname = reader[data["lastnameColName"]].ToString();
                        string newLine = string.Format("{0};{1};{2};{3}", birthdate.ToShortDateString(), lastname, firstname, age);
                        Console.WriteLine(newLine);
                        csv.AppendLine(newLine);
                    }
                    reader.Close();
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }
            
            ShowCSVSaveDialog(csv, data["csvDefaultNameBirthdate"], data["csvInitialDirBirthdate"]);
        }
        private static void CreateBloodCSV()
        {
            var csv = new StringBuilder();
            csv.AppendLine(data["csvHeaderBlood"]);

            var queryString = data["bloodQuery"];
            using (OleDbConnection connection = new OleDbConnection(data["connectionString"]))
            using (OleDbCommand command = new OleDbCommand(queryString, connection))
            {
                try
                {
                    connection.Open();
                    OleDbDataReader reader = command.ExecuteReader();
                    while (reader.Read())
                    {
                        var desc = reader[data["descriptionColName"]];
                        var bloodTime = ((DateTime)reader[data["timeColName"]]);
                        var firstname = reader[data["firstnameColName"]].ToString();
                        var lastname = reader[data["lastnameColName"]].ToString();

                        string newLine = string.Format("{0};{1};{2};{3}", desc, bloodTime.ToShortDateString(), lastname, firstname);
                        Console.WriteLine(newLine);
                        csv.AppendLine(newLine);
                    }
                    reader.Close();
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }

            ShowCSVSaveDialog(csv, data["csvDefaultNameBlood"], data["csvInitialDirBlood"]);
        }

        private static void ShowCSVSaveDialog(StringBuilder csv, string defaultName, string initialDir)
        {

            using (var dialog = new SaveFileDialog())
            {
                dialog.FileName = string.Format(defaultName, months[DateTime.Now.Month - 1]);
                dialog.InitialDirectory = initialDir;
                dialog.OverwritePrompt = true;

                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    File.WriteAllText(dialog.FileName, csv.ToString());
                }
            }
        }

        private static bool ReadDataFromFile(string filename)
        {
            if(!File.Exists(filename))
            {
                Console.WriteLine("Keine Datei db.ini gefunden, diese sollte im Ordner des Kasterls (Anwendung) liegen.");
                return false;
            }
            foreach (var row in File.ReadAllLines(filename))
                data.Add(row.Split('=')[0], string.Join("=", row.Split('=').Skip(1).ToArray()));
            return true;
        }
    }
}
