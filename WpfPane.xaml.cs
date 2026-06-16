using ExcelDna.Integration;

namespace ExcelDnaPoc;

// Contenu WPF du volet. Interagit avec Excel via ExcelDnaUtil.Application.
// Base qualifiee en entier : ImplicitUsings importe aussi System.Windows.Forms,
// donc "UserControl" seul serait ambigu (WinForms vs WPF).
public partial class WpfPane : System.Windows.Controls.UserControl
{
    public WpfPane()
    {
        InitializeComponent();
        Log("Volet WPF charge.");
    }

    // Ajoute une ligne d'UI pour un traitement et branche les callbacks <-> JobRow.
    // Plusieurs jobs coexistent => plusieurs lignes (barres) en parallele.
    public void AddJob(JokeJob job, string label)
    {
        var row = new JobRow();
        row.SetStatus($"{label} : demarrage...");
        row.CancelRequested += job.Cancel;

        job.OnStatus = s => row.SetStatus($"{label} : {s}");
        job.OnProgress = row.SetProgress;
        job.OnBusy = row.SetBusy;
        job.OnDone = () => { row.Finish(); RemoveLater(row); };

        JobsPanel.Children.Add(row);
    }

    // Retire la ligne ~4s apres la fin (laisse voir l'etat final).
    private void RemoveLater(JobRow row) => OnUi(async () =>
    {
        await System.Threading.Tasks.Task.Delay(4000);
        JobsPanel.Children.Remove(row);
    });

    private void BtnWrite_Click(object sender, System.Windows.RoutedEventArgs e)
    {
        try
        {
            ((dynamic)ExcelDnaUtil.Application).ActiveCell.Value2 = InputBox.Text;
            Log($"Ecrit dans la cellule active : \"{InputBox.Text}\"");
        }
        catch (System.Exception ex) { Log("Erreur : " + ex.Message); }
    }

    private void BtnRead_Click(object sender, System.Windows.RoutedEventArgs e)
    {
        try
        {
            dynamic cell = ((dynamic)ExcelDnaUtil.Application).ActiveCell;
            object? val = cell.Value2;
            string addr = cell.Address;
            Log($"Cellule {addr} = {(val ?? "(vide)")}");
        }
        catch (System.Exception ex) { Log("Erreur : " + ex.Message); }
    }

    private void BtnCount_Click(object sender, System.Windows.RoutedEventArgs e)
    {
        try
        {
            dynamic selection = ((dynamic)ExcelDnaUtil.Application).Selection;
            int count = 0;
            foreach (dynamic cell in selection.Cells)
                if (cell.Value2 != null) count++;
            Log($"Cellules non vides dans la selection : {count}");
        }
        catch (System.Exception ex) { Log("Erreur : " + ex.Message); }
    }

    private void Log(string message)
    {
        LogBox.Text = $"{System.DateTime.Now:HH:mm:ss}  {message}\r\n" + LogBox.Text;
    }

    // Execute l'action sur le thread UI WPF (sans bloquer l'appelant).
    private void OnUi(System.Action action)
    {
        if (Dispatcher.CheckAccess()) action();
        else Dispatcher.BeginInvoke(action);
    }
}
