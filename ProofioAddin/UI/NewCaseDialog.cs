using ProofioAddIn.Api;
using System;
using System.Drawing;
using System.Windows.Forms;

namespace ProofioAddIn.UI
{
    public sealed class NewCaseDialog : Form
    {
        public CreateCaseRequest Request { get; private set; }

        private readonly TextBox  _titleBox;
        private readonly ComboBox _artBox;
        private readonly TextBox  _vornameBox;
        private readonly TextBox  _nachnameBox;
        private readonly TextBox  _firmaBox;
        private readonly TextBox  _emailBox;

        public NewCaseDialog()
        {
            Text            = "Neuen Fall anlegen";
            StartPosition   = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MinimizeBox     = false;
            MaximizeBox     = false;
            ClientSize      = new Size(500, 540);
            BackColor       = UITheme.Background;
            Font            = UITheme.DefaultFont;
            Icon            = UITheme.AppIcon;

            // ── Title bar ─────────────────────────────────────────────────
            var titleBar = new ProofioTitleBar("Neuen Fall anlegen") { Dock = DockStyle.Top };

            // ── Footer ────────────────────────────────────────────────────
            var footer = new Panel { Dock = DockStyle.Bottom, Height = 56, BackColor = UITheme.Surface };
            footer.Paint += (s, e) => e.Graphics.DrawLine(new Pen(UITheme.Border), 0, 0, footer.Width, 0);

            var btnCancel = new StyledButton("Abbrechen",  false) { DialogResult = DialogResult.Cancel };
            var btnOk     = new StyledButton("Fall anlegen", true);
            btnOk.Click  += Ok_Click;

            footer.Layout += (s, e) =>
            {
                int cy = footer.Height / 2;
                btnCancel.Location = new Point(footer.Width - btnCancel.Width - btnOk.Width - 12, cy - btnCancel.Height / 2);
                btnOk.Location     = new Point(footer.Width - btnOk.Width - 4,                    cy - btnOk.Height / 2);
            };
            footer.Controls.AddRange(new Control[] { btnCancel, btnOk });

            // ── Body ──────────────────────────────────────────────────────
            var body = new Panel { Dock = DockStyle.Fill, BackColor = UITheme.Background, Padding = new Padding(20, 16, 20, 8) };

            // Use a TableLayoutPanel for clean label + field alignment
            var tbl = new TableLayoutPanel
            {
                Dock        = DockStyle.Fill,
                ColumnCount = 2,
                RowCount    = 8,
                BackColor   = UITheme.Background,
                AutoSize    = false
            };
            tbl.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 130));
            tbl.ColumnStyles.Add(new ColumnStyle(SizeType.Percent,  100));
            for (int i = 0; i < 8; i++)
                tbl.RowStyles.Add(new RowStyle(SizeType.Absolute, i == 1 ? 30 : 56)); // row 1 = section header

            int row = 0;

            // Titel
            tbl.Controls.Add(MakeLabel("Titel *"), 0, row);
            _titleBox = MakeField();
            tbl.Controls.Add(_titleBox, 1, row); row++;

            // Section header
            var sectionLbl = new Label
            {
                Text      = "AUFTRAGGEBER  (optional)",
                Font      = UITheme.LabelFont,
                ForeColor = UITheme.TextMuted,
                Dock      = DockStyle.Fill,
                TextAlign = ContentAlignment.BottomLeft
            };
            tbl.SetColumnSpan(sectionLbl, 2);
            tbl.Controls.Add(sectionLbl, 0, row); row++;

            // Vorname
            tbl.Controls.Add(MakeLabel("Vorname"), 0, row);
            _vornameBox = MakeField();
            tbl.Controls.Add(_vornameBox, 1, row); row++;

            // Nachname
            tbl.Controls.Add(MakeLabel("Nachname"), 0, row);
            _nachnameBox = MakeField();
            tbl.Controls.Add(_nachnameBox, 1, row); row++;

            // Firma
            tbl.Controls.Add(MakeLabel("Firma"), 0, row);
            _firmaBox = MakeField();
            tbl.Controls.Add(_firmaBox, 1, row); row++;

            // E-Mail
            tbl.Controls.Add(MakeLabel("E-Mail"), 0, row);
            _emailBox = MakeField();
            tbl.Controls.Add(_emailBox, 1, row); row++;

            // Art
            tbl.Controls.Add(MakeLabel("Art"), 0, row);
            _artBox = new ComboBox
            {
                Dock          = DockStyle.Fill,
                DropDownStyle = ComboBoxStyle.DropDownList,
                FlatStyle     = FlatStyle.Flat,
                Font          = UITheme.DefaultFont,
                BackColor     = Color.White,
                ForeColor     = UITheme.TextPrimary,
                Margin        = new Padding(0, 8, 0, 0)
            };
            _artBox.Items.AddRange(new object[] { "privatgutachten", "gerichtsgutachten", "schiedsgutachten", "parteigutachten" });
            _artBox.SelectedIndex = 0;
            tbl.Controls.Add(_artBox, 1, row); row++;

            // Hint
            var hint = new Label
            {
                Text      = "* Pflichtfeld",
                Font      = UITheme.SmallFont,
                ForeColor = UITheme.TextMuted,
                Dock      = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft
            };
            tbl.SetColumnSpan(hint, 2);
            tbl.Controls.Add(hint, 0, row);

            body.Controls.Add(tbl);

            Controls.Add(body);
            Controls.Add(footer);
            Controls.Add(titleBar);

            AcceptButton = btnOk;
            CancelButton = btnCancel;
        }

        private void Ok_Click(object sender, EventArgs e)
        {
            var title = (_titleBox.Text ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(title))
            {
                MessageBox.Show("Bitte geben Sie einen Titel ein.", "Proofio", MessageBoxButtons.OK, MessageBoxIcon.Information);
                DialogResult = DialogResult.None;
                _titleBox.Focus();
                return;
            }
            Request = new CreateCaseRequest
            {
                Title                = title,
                Art                  = _artBox.SelectedItem?.ToString(),
                AuftraggeberVorname  = NullIfEmpty(_vornameBox.Text),
                AuftraggeberNachname = NullIfEmpty(_nachnameBox.Text),
                AuftraggeberFirma    = NullIfEmpty(_firmaBox.Text),
                AuftraggeberEmail    = NullIfEmpty(_emailBox.Text)
            };
            DialogResult = DialogResult.OK;  // closes the dialog
        }

        private static Label MakeLabel(string text) => new Label
        {
            Text      = text,
            Dock      = DockStyle.Fill,
            Font      = UITheme.LabelFont,
            ForeColor = UITheme.TextMuted,
            TextAlign = ContentAlignment.MiddleLeft
        };

        private static TextBox MakeField() => new TextBox
        {
            Dock        = DockStyle.Fill,
            BorderStyle = BorderStyle.FixedSingle,
            Font        = UITheme.DefaultFont,
            ForeColor   = UITheme.TextPrimary,
            BackColor   = Color.White,
            Margin      = new Padding(0, 8, 0, 0)
        };

        private static string NullIfEmpty(string s) =>
            string.IsNullOrWhiteSpace(s) ? null : s.Trim();
    }
}
