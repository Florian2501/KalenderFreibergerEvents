using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Ical.Net;
using Ical.Net.DataTypes;
using Ical.Net.Serialization;
using System.Net.Mail;
using System.IO;
using Ical.Net.CalendarComponents;
using System.Net;
using System;
using System.IO;
using System.Windows;
using Microsoft.Win32;

namespace CalenderTest
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }


        public void FillCalendar(Calendar calendar, string path)
        {
            try
            {
                using (StreamReader sr = new StreamReader(path, Encoding.UTF8))
                {
                    //if the date file exists open it
                    string line = "";
                    //Skip first line because its the header    
                    sr.ReadLine();
                    //go through the date file line by line
                    while ((line = sr.ReadLine()) != null)
                    {
                        //split the line into the wanted pieces
                        string[] eventInfos = line.Split('\t');

                        //check wether this line is empty and got to the next one if yes
                        if (string.IsNullOrEmpty(eventInfos[3])) continue;

                        DateTime date;
                        DateTime start;
                        DateTime end;
                        try
                        {
                            date = Convert.ToDateTime(eventInfos[1]);//je immer die x.te Spalte in der Tabelle als Ziffer eintragen
                            start = Convert.ToDateTime(eventInfos[5]);
                            end = Convert.ToDateTime(eventInfos[6]);

                        }
                        catch (FormatException)
                        {
                            MessageBox.Show($"Problems in the Date file. A value could not be converted correctly.\nIt will not count.");
                            continue;

                        }
                        catch (OverflowException)
                        {
                            MessageBox.Show($"Problems in the Date file. Value was to big.\nIt will not count.");
                            continue;
                        }
                        catch (IndexOutOfRangeException)
                        {
                            MessageBox.Show($"Problems in the Date file. Index out of range.\nIt will not count.");
                            continue;
                        }

                        calendar.Events.Add(new CalendarEvent
                        {
                            Class = "PUBLIC",
                            Summary = eventInfos[3] + ": " + eventInfos[4],
                            Created = new CalDateTime(DateTime.Now),
                            Description = eventInfos[7] + "\nBemerkung: " + eventInfos[8],
                            //Erstelle die Anfangszeit mithilfe des Datumseintrags und die Uhrzeit aus den Von und Bis Spalten, 0 steht für Sekunden
                            Start = new CalDateTime(new DateTime(date.Year, date.Month, date.Day, start.Hour, start.Minute, 0)),
                            End = new CalDateTime(new DateTime(date.Year, date.Month, date.Day, end.Hour, end.Minute, 0)),
                            Sequence = 0,
                            Uid = Guid.NewGuid().ToString(),
                            Location = eventInfos[9]
                        });
                    }
                }
            }

            catch (IOException)
            {
                MessageBox.Show("The Date file could not be read. It probably does not exist in the given location."); 
            }
            catch (Exception)
            {
                MessageBox.Show("Something with the Date file went wrong.");
            }
        }
        
        private void Button_Click(object sender, RoutedEventArgs e)
        {

            string inputPath = @"./in.tsv";
            string outputPath = @"./FreibergerEvents.ics";

            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Tab Separated Values-Datei (.tsv)|*.tsv";
            if(openFileDialog.ShowDialog() == true)
            {
                inputPath = openFileDialog.FileName;
            }

            Calendar calendar = new Calendar();

            FillCalendar(calendar, inputPath);

            MessageBox.Show("Bitte gib als nächstes  den Ort an, wo die Datei gespeichert werden soll.\nWenn du einfach direkt auf Speichern klickst, wird die Datei unter dem Namen \"FreibergerEvents.ics\" an die gleiche Stelle wie deine Eingabedatei gespeichert.");

            openFileDialog.Filter = "Kalenderdateien (.ics)|*.ics";
            openFileDialog.ValidateNames = false;
            openFileDialog.CheckFileExists = false;
            openFileDialog.CheckPathExists = true;
            string[] tmp = openFileDialog.FileName.Split(@"\");
            tmp[openFileDialog.FileName.Split(@"\").Length - 1] = "FreibergerEvents.ics";
            openFileDialog.FileName = string.Join(@"\",tmp);

            if (openFileDialog.ShowDialog() == true)
            {
                outputPath = openFileDialog.FileName;
            }
            else 
            {
                MessageBox.Show("Speichern abgebrochen! Keine Angst, du wurdest nicht gehackt und es wurde keine Datei erzeugt :D");
                return;
            }

            using (StreamWriter sw = new StreamWriter(outputPath))
            {
                sw.AutoFlush = true;
                var serializer = new CalendarSerializer(new SerializationContext());
                var serializedCalendar = serializer.SerializeToString(calendar);
                sw.WriteLine(serializedCalendar);
            }

            MessageBox.Show("Die Kalenderdatei wurde erstellt und im angegebenen Verzeichnis abgespeichert.\nViel Spaß bei den Veranstaltungen :)");
        }
    }
}
