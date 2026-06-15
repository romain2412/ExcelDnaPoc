using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Windows.Forms.Integration;

namespace ExcelDnaPoc;

// Le Custom Task Pane d'Office/ExcelDna ne peut heberger qu'un controle ActiveX/WinForms
// COM-visible. On expose donc ce UserControl WinForms qui contient un ElementHost,
// lequel embarque le vrai contenu WPF (WpfPane).
[ComVisible(true)]
[Guid("B5E3F1A2-7C44-4E9A-9D2B-1F3A6C8D0E11")]
[ClassInterface(ClassInterfaceType.None)]
public class WpfPaneHost : UserControl
{
    // Instance du contenu WPF, exposee pour que le ruban puisse piloter
    // le statut et la barre de progression.
    public WpfPane Pane { get; }

    public WpfPaneHost()
    {
        Pane = new WpfPane();
        var elementHost = new ElementHost
        {
            Dock = DockStyle.Fill,
            Child = Pane,
        };
        Controls.Add(elementHost);
    }
}
