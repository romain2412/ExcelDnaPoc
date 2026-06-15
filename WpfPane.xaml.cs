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

    // ----- API publique pilotee par l'appel async du ruban -----
    // Ces methodes peuvent etre appelees depuis un thread de fond : on marshale
    // systematiquement vers le thread UI WPF via le Dispatcher.

    public void SetStatus(string message) => OnUi(() => StatusText.Text = message);

    public void SetProgress(double percent) => OnUi(() =>
    {
        Progress.IsIndeterminate = false;
        Progress.Value = percent;
    });

    // Mode "occupe" sans pourcentage (ex. pendant l'appel reseau de duree inconnue).
    public void SetBusy(bool busy) => OnUi(() => Progress.IsIndeterminate = busy);

    // Execute l'action sur le thread UI WPF (sans bloquer l'appelant).
    private void OnUi(System.Action action)
    {
        if (Dispatcher.CheckAccess()) action();
        else Dispatcher.BeginInvoke(action);
    }
}
