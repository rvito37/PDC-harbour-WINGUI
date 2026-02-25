namespace PdcGui;

static class Program
{
    [STAThread]
    static void Main(string[] args)
    {
        ApplicationConfiguration.Initialize();

        // Username: from command line arg, or prompt
        string userName = args.Length > 0 ? args[0] : PromptUserName();

        if (string.IsNullOrWhiteSpace(userName))
        {
            MessageBox.Show("No username provided. Exiting.", "PDC",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        Application.Run(new MainForm(userName));
    }

    private static string PromptUserName()
    {
        // Simple login dialog
        using var dlg = new Form
        {
            Text = "PDC Login",
            Size = new Size(350, 180),
            StartPosition = FormStartPosition.CenterScreen,
            FormBorderStyle = FormBorderStyle.FixedDialog,
            MaximizeBox = false,
            MinimizeBox = false,
            Font = new Font("Segoe UI", 10f)
        };

        var lbl = new Label
        {
            Text = "User Name:",
            Location = new Point(20, 25),
            AutoSize = true
        };

        var txt = new TextBox
        {
            Location = new Point(120, 22),
            Size = new Size(180, 28),
            Text = Environment.UserName
        };

        var btnOk = new Button
        {
            Text = "Login",
            DialogResult = DialogResult.OK,
            Location = new Point(120, 65),
            Size = new Size(80, 32)
        };

        var btnCancel = new Button
        {
            Text = "Cancel",
            DialogResult = DialogResult.Cancel,
            Location = new Point(210, 65),
            Size = new Size(80, 32)
        };

        dlg.Controls.AddRange(new Control[] { lbl, txt, btnOk, btnCancel });
        dlg.AcceptButton = btnOk;
        dlg.CancelButton = btnCancel;

        return dlg.ShowDialog() == DialogResult.OK ? txt.Text.Trim() : "";
    }
}
