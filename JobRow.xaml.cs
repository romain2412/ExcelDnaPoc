using System;

namespace ExcelDnaPoc;

// Une ligne d'UI representant UN traitement : statut + barre de progression + Annuler.
// Methodes thread-safe (marshalent vers le thread UI WPF). Le clic sur Annuler leve
// CancelRequested (que WpfPane branche sur JokeJob.Cancel).
public partial class JobRow : System.Windows.Controls.UserControl
{
    public event Action? CancelRequested;

    public JobRow() => InitializeComponent();

    private void BtnCancel_Click(object sender, System.Windows.RoutedEventArgs e)
        => CancelRequested?.Invoke();

    public void SetStatus(string text) => OnUi(() => StatusText.Text = text);

    public void SetProgress(double percent) => OnUi(() =>
    {
        Progress.IsIndeterminate = false;
        Progress.Value = percent;
    });

    public void SetBusy(bool busy) => OnUi(() => Progress.IsIndeterminate = busy);

    // Fin du traitement : on desactive le bouton Annuler.
    public void Finish() => OnUi(() => BtnCancel.IsEnabled = false);

    private void OnUi(Action action)
    {
        if (Dispatcher.CheckAccess()) action();
        else Dispatcher.BeginInvoke(action);
    }
}
