using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Timers;
using System.IO;
using ClosedXML.Excel;

namespace ArbeitszeitErfassung;

/// <summary>
/// Interaction logic for MainWindow.xaml
/// </summary>
public partial class MainWindow : Window
{
    static DateTime startTime;
    static DateTime pauseStartTime;
    static TimeSpan pauseTime;
    static TimeSpan pauseTimeCache;
    static TimeSpan workTime;
    private static System.Timers.Timer startClock;

    private bool isStarted = false;
    private bool pauseIsStarted = false;

    string dateiNameCSV;
    string dateiNameExcel;

    public MainWindow()
    {
        InitializeComponent();

        if (!string.IsNullOrWhiteSpace(Properties.Settings.Default.Speicherpfad))
        {
            Text1.Text = "Gespeicherter Pfad: " + Properties.Settings.Default.Speicherpfad;
        }
    }

    private void Bestätigen_Click(object sender, RoutedEventArgs e)
    {
        string eingabePfad = BenutzereingabeBox.Text;

        if (Directory.Exists(eingabePfad))
        {
            Properties.Settings.Default.Speicherpfad = eingabePfad;
            Properties.Settings.Default.Save();
            MessageBox.Show("Pfad gespeichert: " + eingabePfad);
        }
        else
        {
            MessageBox.Show("Pfad ist ungültig oder existiert nicht.");
        }
    }

    private void Button_Click(object sender, RoutedEventArgs e)
    {
        if (isStarted == false)
        {
            SetTimer();
            isStarted = true;
            startTime = DateTime.Now;
            Text1.Text = "Start Zeit: " + startTime;
            Button1.Content = "Zeiterfassung gestartet";
        }
        else
        {
            var result = MessageBox.Show("Erfassung läuft bereits. " +
                "Möchtest du sie neu starten?", "Warnung", MessageBoxButton.YesNo, MessageBoxImage.Information);

            switch (result)
            {
                case MessageBoxResult.Cancel:
                    
                    break;
                case MessageBoxResult.Yes:
                    startTime = DateTime.Now;
                    Text1.Text = "Start Zeit: " + startTime;
                    Button1.Content = "Zeiterfassung gestartet";
                    if (pauseTime != pauseTimeCache)
                    {
                        pauseTime = pauseTime - pauseTime;
                        Text5.Text = "Pausenzeit Insgesamt: " + $"{(int)pauseTime.TotalHours:D2}:{pauseTime.Minutes:D2}:{pauseTime.Seconds:D2}";
                    }
                    break;
                case MessageBoxResult.No:
                    
                    break;
            }
        }
    }

    private void Button_Click2(object sender, RoutedEventArgs e)
    {
        if (!isStarted) return;

        string temp = Button3.Content.ToString();
        if (temp == "Pause Starten")
        {
            pauseIsStarted = true;
            pauseStartTime = DateTime.Now;
            Text2.Text = "Pause gestartet um: " + pauseStartTime;
            Button3.Content = "Pause Beenden";
        }
        else
        {
            pauseIsStarted = false;
            TimeSpan elapsed = DateTime.Now - pauseStartTime;
            pauseTime += elapsed;
            Button3.Content = "Pause Starten";
            Text5.Text = "Pausenzeit Insgesamt: " + $"{(int)pauseTime.TotalHours:D2}:{pauseTime.Minutes:D2}:{pauseTime.Seconds:D2}";
            Text3.Text = "";
        }
    }

    private void SetTimer()
    {
        startClock = new System.Timers.Timer(500);
        startClock.Elapsed += OnTimedEvent;
        startClock.AutoReset = true;
        startClock.Enabled = true;
    }

    private void OnTimedEvent(object sender, ElapsedEventArgs e)
    {
        Dispatcher.Invoke(() =>
        {
            if (isStarted && !pauseIsStarted)
            {
                TimeSpan elapsed = DateTime.Now - startTime - pauseTime;
                workTime = elapsed;
                Text4.Text = "Verstrichene Zeit: " + $"{(int)elapsed.TotalHours:D2}:{elapsed.Minutes:D2}:{elapsed.Seconds:D2}";
            }
            else if (isStarted && pauseIsStarted)
            {
                TimeSpan currentPause = DateTime.Now - pauseStartTime;
                Text3.Text = "Verstrichene Pausenzeit: " + $"{(int)currentPause.TotalHours:D2}:{currentPause.Minutes:D2}:{currentPause.Seconds:D2}";
            }
        });
    }

    private void Save(object sender, RoutedEventArgs e)
    {
        string pfad = Properties.Settings.Default.Speicherpfad;

        if (string.IsNullOrWhiteSpace(pfad) || !Directory.Exists(pfad))
        {
            MessageBox.Show("Kein gültiger Speicherpfad vorhanden. Bitte zuerst bestätigen.");
            return;
        }
        dateiNameCSV = "ArbeitszeitenCSV-" + startTime.Year + "-" + startTime.Month + ".csv";
        string csvPath = System.IO.Path.Combine(pfad, dateiNameCSV);

        using (StreamWriter datei = new StreamWriter(csvPath, true))
        {
            datei.WriteLine("Start: ;" + startTime);
            datei.WriteLine("Ende: ;" + DateTime.Now);
            datei.WriteLine("Dauer der Pausen: ;" + pauseTime.ToString(@"hh\:mm\:ss"));
            datei.WriteLine("Arbeitszeit: ;" + workTime.ToString(@"hh\:mm\:ss"));
            datei.WriteLine("___________________;_____________________");
        }

        SaveToExcel(pfad);
        MessageBox.Show("Datei Erfolgreich gespeichert gespeichert unter: " + pfad,"Erfolgreich Gespeichert");
    }

    private void SaveToExcel(string pfad)
    {
        dateiNameExcel = "ArbeitszeitenExcel-" + startTime.Year + "-" + startTime.Month + ".xlsx";
        string path = System.IO.Path.Combine(pfad, dateiNameExcel);

        XLWorkbook workbook;
        IXLWorksheet worksheet;

        if (File.Exists(path))
        {
            workbook = new XLWorkbook(path);
            worksheet = workbook.Worksheet(1);
        }
        else
        {
            workbook = new XLWorkbook();
            worksheet = workbook.Worksheets.Add("Arbeitszeiten");
        }

        int lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 0;
        int newRow = lastRow + 2;

        worksheet.Cell(newRow, 1).Value = "Start:";
        worksheet.Cell(newRow, 2).Value = startTime.ToString("dd.MM.yyyy HH:mm");

        worksheet.Cell(newRow + 1, 1).Value = "Ende:";
        worksheet.Cell(newRow + 1, 2).Value = DateTime.Now.ToString("dd.MM.yyyy HH:mm");

        worksheet.Cell(newRow + 2, 1).Value = "Dauer der Pausen:";
        worksheet.Cell(newRow + 2, 2).Value = pauseTime.ToString(@"hh\:mm\:ss");

        TimeSpan berechneteArbeitszeit = DateTime.Now - startTime - pauseTime;
        worksheet.Cell(newRow + 3, 1).Value = "Arbeitszeit:";
        worksheet.Cell(newRow + 3, 2).Value = berechneteArbeitszeit.ToString(@"hh\:mm\:ss");

        worksheet.Range(newRow, 1, newRow + 3, 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
        worksheet.Columns().AdjustToContents();

        workbook.SaveAs(path);
    }
}