using System.Windows.Forms;

namespace ExcelWorkbookMerger.Models;

public class ErrorDialog : Form
{
    public ErrorDialog(string message)
    {
        var textBox = new TextBox
        {
            Text = message,
            ReadOnly = true,
            Multiline = true,
            Dock = DockStyle.Fill,
            ScrollBars = ScrollBars.Vertical
        };
        var button = new Button
        {
            Text = "OK",
            Dock = DockStyle.Bottom,
            DialogResult = DialogResult.OK
        };
        AcceptButton = button;
        Controls.Add(textBox);
        Controls.Add(button);
        Text = "Error";
        Width = 500;
        Height = 300;
    }
}