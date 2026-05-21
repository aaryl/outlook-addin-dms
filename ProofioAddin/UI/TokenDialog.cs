using ProofioAddIn.Services;
using System;
using System.Drawing;
using System.Windows.Forms;

namespace ProofioAddIn.UI
{
    public sealed class TokenDialog : Form
    {
        private readonly TokenStore _tokenStore;
        private readonly TextBox    _tokenBox;
        private readonly TextBox    _apiUrlBox;
        private readonly ComboBox   _sendModeBox;

        public TokenDialog(TokenStore tokenStore)
        {
            _tokenStore = tokenStore ?? throw new ArgumentNullException(nameof(tokenStore));

            Text            = "Proofio – Einstellungen";
            StartPosition   = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MinimizeBox     = false;
            MaximizeBox     = false;
            ClientSize      = new Size(540, 310);
            BackColor       = UITheme.Background;
            Font            = UITheme.DefaultFont;
            Icon            = UITheme.AppIcon;

            // ── Title bar ─────────────────────────────────────────────────
            var titleBar = new ProofioTitleBar("Einstellungen") { Dock = DockStyle.Top };

            // ── Footer ────────────────────────────────────────────────────
            var footer = new Panel { Dock = DockStyle.Bottom, Height = 56, BackColor = UITheme.Surface };
            footer.Paint += (s, e) => e.Graphics.DrawLine(new Pen(UITheme.Border), 0, 0, footer.Width, 0);

            var btnCancel = new StyledButton("Abbrechen", false) { DialogResult = DialogResult.Cancel };
            var btnSave   = new StyledButton("Speichern", true);
            btnSave.Click += Save_Click;

            footer.Layout += (s, e) =>
            {
                int cy = footer.Height / 2;
                btnCancel.Location = new Point(footer.Width - btnCancel.Width - btnSave.Width - 12, cy - btnCancel.Height / 2);
                btnSave.Location   = new Point(footer.Width - btnSave.Width - 4,                    cy - btnSave.Height / 2);
            };
            footer.Controls.AddRange(new Control[] { btnCancel, btnSave });

            // ── Body ──────────────────────────────────────────────────────
            var body = new Panel { Dock = DockStyle.Fill, BackColor = UITheme.Background, Padding = new Padding(20, 16, 20, 8) };

            var tbl = new TableLayoutPanel
            {
                Dock        = DockStyle.Fill,
                ColumnCount = 2,
                RowCount    = 4,
                BackColor   = UITheme.Background
            };
            tbl.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 130));
            tbl.ColumnStyles.Add(new ColumnStyle(SizeType.Percent,  100));
            tbl.RowStyles.Add(new RowStyle(SizeType.Absolute, 46)); // token
            tbl.RowStyles.Add(new RowStyle(SizeType.Absolute, 46)); // api url
            tbl.RowStyles.Add(new RowStyle(SizeType.Absolute, 46)); // send mode
            tbl.RowStyles.Add(new RowStyle(SizeType.Percent,  100)); // hint

            // Bearer-Token
            tbl.Controls.Add(MakeLabel("Bearer-Token"), 0, 0);
            _tokenBox = new TextBox
            {
                Dock                  = DockStyle.Fill,
                BorderStyle           = BorderStyle.FixedSingle,
                Font                  = new Font("Courier New", 9f),
                ForeColor             = UITheme.TextPrimary,
                BackColor             = Color.White,
                UseSystemPasswordChar = true,
                Text                  = _tokenStore.GetToken() ?? string.Empty,
                Margin                = new Padding(0, 8, 0, 0)
            };
            tbl.Controls.Add(_tokenBox, 1, 0);

            // API-URL
            tbl.Controls.Add(MakeLabel("API-Basis-URL"), 0, 1);
            _apiUrlBox = new TextBox
            {
                Dock        = DockStyle.Fill,
                BorderStyle = BorderStyle.FixedSingle,
                Font        = UITheme.DefaultFont,
                ForeColor   = UITheme.TextPrimary,
                BackColor   = Color.White,
                Text        = _tokenStore.GetApiBaseUrl(),
                Margin      = new Padding(0, 8, 0, 0)
            };
            tbl.Controls.Add(_apiUrlBox, 1, 1);

            // Beim Senden
            tbl.Controls.Add(MakeLabel("Beim Senden"), 0, 2);
            _sendModeBox = new ComboBox
            {
                Dock          = DockStyle.Fill,
                DropDownStyle = ComboBoxStyle.DropDownList,
                FlatStyle     = FlatStyle.Flat,
                Font          = UITheme.DefaultFont,
                BackColor     = Color.White,
                ForeColor     = UITheme.TextPrimary,
                Margin        = new Padding(0, 8, 0, 0)
            };
            _sendModeBox.Items.Add(new ModeItem("Nie automatisch ablegen",           TokenStore.SendFilingModeNever));
            _sendModeBox.Items.Add(new ModeItem("Immer Akte auswählen und ablegen",  TokenStore.SendFilingModeAlways));
            _sendModeBox.Items.Add(new ModeItem("Vor dem Senden fragen",             TokenStore.SendFilingModeAsk));
            SelectMode(_tokenStore.GetSendFilingMode());
            tbl.Controls.Add(_sendModeBox, 1, 2);

            // Hint
            var hint = new Panel { Dock = DockStyle.Fill, BackColor = UITheme.AccentLight, Margin = new Padding(0, 8, 0, 0) };
            hint.Paint += (s, e) => e.Graphics.DrawRectangle(new Pen(Color.FromArgb(190, 205, 240)), 0, 0, hint.Width - 1, hint.Height - 1);
            var hintLbl = new Label
            {
                Text      = "Token in Proofio kopieren:  Outlook-Add-in → Token kopieren",
                Dock      = DockStyle.Fill,
                Font      = UITheme.SmallFont,
                ForeColor = UITheme.TextSecondary,
                TextAlign = ContentAlignment.MiddleLeft,
                Padding   = new Padding(10, 0, 0, 0)
            };
            hint.Controls.Add(hintLbl);
            tbl.SetColumnSpan(hint, 2);
            tbl.Controls.Add(hint, 0, 3);

            body.Controls.Add(tbl);

            Controls.Add(body);
            Controls.Add(footer);
            Controls.Add(titleBar);

            AcceptButton = btnSave;
            CancelButton = btnCancel;
        }

        private void Save_Click(object sender, EventArgs e)
        {
            try
            {
                _tokenStore.Save(_tokenBox.Text.Trim(), _apiUrlBox.Text.Trim(), GetSelectedMode());
                MessageBox.Show("Proofio-Einstellungen wurden gespeichert.", "Proofio", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                Logger.Error("Token speichern fehlgeschlagen.", ex);
                MessageBox.Show("Die Einstellungen konnten nicht gespeichert werden: " + ex.Message, "Proofio", MessageBoxButtons.OK, MessageBoxIcon.Error);
                DialogResult = DialogResult.None;
            }
        }

        private string GetSelectedMode() =>
            (_sendModeBox.SelectedItem as ModeItem)?.Value ?? TokenStore.SendFilingModeNever;

        private void SelectMode(string value)
        {
            for (int i = 0; i < _sendModeBox.Items.Count; i++)
            {
                var item = _sendModeBox.Items[i] as ModeItem;
                if (item != null && string.Equals(item.Value, value, StringComparison.OrdinalIgnoreCase))
                { _sendModeBox.SelectedIndex = i; return; }
            }
            _sendModeBox.SelectedIndex = 0;
        }

        private static Label MakeLabel(string text) => new Label
        {
            Text      = text,
            Dock      = DockStyle.Fill,
            Font      = UITheme.LabelFont,
            ForeColor = UITheme.TextMuted,
            TextAlign = ContentAlignment.MiddleLeft
        };

        private sealed class ModeItem
        {
            public ModeItem(string text, string value) { Text = text; Value = value; }
            public string Text  { get; }
            public string Value { get; }
            public override string ToString() => Text;
        }
    }
}
