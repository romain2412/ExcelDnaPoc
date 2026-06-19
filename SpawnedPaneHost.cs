using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace ExcelDnaPoc;

// Hote LEGER d'un volet "spawn dynamique" : juste un libelle qui identifie le
// volet (numero + cote d'ancrage). Chaque volet cree par MultiPaneManager possede
// SA PROPRE instance de cet UserControl -> demontre que N volets coexistent, chacun
// avec son contenu. (Le volet WPF riche, lui, vit dans WpfPaneHost/WpfPane.)
//
// ComVisible + Guid : un Custom Task Pane d'Office heberge un controle COM ; chaque
// classe hote doit etre COM-visible avec un GUID stable. Plusieurs INSTANCES de la
// meme classe (donc plusieurs volets) sont parfaitement valides.
[ComVisible(true)]
[Guid("C7A1D2E3-4B56-4789-9A0B-2C3D4E5F6071")]
[ClassInterface(ClassInterfaceType.None)]
public class SpawnedPaneHost : UserControl
{
    private readonly Label _label;

    public SpawnedPaneHost()
    {
        _label = new Label
        {
            Dock = DockStyle.Fill,
            TextAlign = ContentAlignment.MiddleCenter,
            Font = new Font("Segoe UI", 12F),
        };
        Controls.Add(_label);
    }

    public void SetInfo(string title, string dock)
        => _label.Text = $"{title}\r\n\r\nAncre a : {dock}\r\n\r\n"
                       + "(Fermez-moi via la croix\r\nou le bouton \"Tout fermer\")";
}
