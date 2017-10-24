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
                //CreateBirthDayCSV();
                CreateBloodCSV();
            }
        }

        private static void CreateBirthDayCSV()
        {
            var csv = new StringBuilder();
            csv.AppendLine(data["csvHeaderBirthdate"]);
            var queryString = data["birthdayQuery"].Replace("$param$", DateTime.Now.Month.ToString());

            using (OleDbConnection connection = new OleDbConnection(data["connectionStringKesztele"]))
            using (OleDbCommand command = new OleDbCommand(queryString, connection))
            {
                try
                {
                    connection.Open();
                    string newLine = "";
                    OleDbDataReader reader = command.ExecuteReader();
                    while (reader.Read())
                    {
                        var birthdate = ((DateTime)reader[data["birthdayColName"]]);
                        var age = DateTime.Now.Year - birthdate.Year;
                        var name = reader[data["bezeichnungColName"]].ToString().Trim();
                        var address = reader[data["addressColName"]].ToString().Trim();
                        var plz = reader[data["PLZColName"]].ToString().Trim();
                        var ort = reader[data["ortColName"]].ToString().Trim();
                        var anrede = reader[data["anredeColName"]].ToString().Trim();
                        var old = newLine;

                        newLine = string.Format("{0};{1};{2};{3};{4};{5}", birthdate.ToShortDateString(), age, name, address, plz, ort, anrede);
                        if(old != newLine)
                        {
                            Console.WriteLine(newLine);
                            csv.AppendLine(newLine);
                        } 
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
            using (OleDbConnection connection = new OleDbConnection(data["connectionStringKesztele"]))
            using (OleDbCommand command = new OleDbCommand(queryString, connection))
            {
                try
                {
                    connection.Open();
                    string newLine = "";
                    OleDbDataReader reader = command.ExecuteReader();
                    while (reader.Read())
                    {
                        var desc = reader[data["descriptionColName"]].ToString().Trim();
                        var bloodTime = ((DateTime)reader[data["timeColName"]]);
                        var name = reader[data["bloodBezeichnungColName"]].ToString().Trim();
                        var old = newLine;

                        newLine = string.Format("{0};{1};{2}", desc, bloodTime.ToShortDateString(), name);
                        if (old != newLine)
                        {
                            Console.WriteLine(newLine);
                            csv.AppendLine(newLine);
                        }
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
                    File.WriteAllText(dialog.FileName, csv.ToString(), Encoding.UTF8);
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
            foreach (var row in File.ReadAllLines(filename, Encoding.UTF8))
            {
                string datum = string.Join("=", row.Split('=').Skip(1).ToArray());
                if (data.ContainsKey("initdir")) //könnte man dynamisch machen :>
                {
                    datum = datum.Replace("$initdir$", data["initdir"]); //Initdir ersetzen!
                }                
                data.Add(row.Split('=')[0], datum);
            }
            return true;
        }
    }
}
