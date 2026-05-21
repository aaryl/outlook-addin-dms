using ProofioAddIn.Api;
using ProofioAddIn.Services;
using System;
using System.Drawing;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ProofioAddIn.UI
{
    public sealed class CasePickerDialog : Form
    {
        private readonly ProofioApiClient _api;
        private readonly TextBox          _searchBox;
        private readonly ListBox          _listBox;
        private readonly Label            _statusLabel;
        private readonly System.Windows.Forms.Timer _debounceTimer;

        public Guid SelectedCaseId { get; private set; }

        public CasePickerDialog(ProofioApiClient api)
        {
            _api = api ?? throw new ArgumentNullException(nameof(api));

            // ── Form ──────────────────────────────────────────────────────
            Text            = "Proofio – Akte auswählen";
            StartPosition   = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.Sizable;
            MinimizeBox     = false;
            MaximizeBox     = false;
            MinimumSize     = new Size(500, 380);
            ClientSize      = new Size(700, 520);
            BackColor       = UITheme.Background;
            Font            = UITheme.DefaultFont;
            Icon            = UITheme.AppIcon;

            // ── Navy title bar (Dock=Top) ──────────────────────────────────
            var titleBar = new ProofioTitleBar("Akte auswählen")
            {
                Dock = DockStyle.Top
            };

            // ── Footer (Dock=Bottom) ───────────────────────────────────────
            var footer = new Panel
            {
                Dock      = DockStyle.Bottom,
                Height    = 56,
                BackColor = UITheme.Surface
            };
            footer.Paint += (s, e) =>
                e.Graphics.DrawLine(new Pen(UITheme.Border), 0, 0, footer.Width, 0);

            _statusLabel = new Label
            {
                AutoSize  = true,
                ForeColor = UITheme.TextMuted,
                Font      = UITheme.SmallFont,
                Dock      = DockStyle.None
            };

            var newCaseLink = new LinkLabel
            {
                Text            = "+ Neuen Fall anlegen",
                Font            = UITheme.SmallBoldFont,
                LinkColor       = UITheme.Accent,
                ActiveLinkColor = UITheme.AccentDark,
                AutoSize        = true,
                Dock            = DockStyle.None
            };
            newCaseLink.LinkClicked += NewCaseLink_Click;

            var btnCancel = new StyledButton("Abbrechen", false) { DialogResult = DialogResult.Cancel };
            var btnSelect = new StyledButton("Auswählen", true)  { DialogResult = DialogResult.OK };
            btnSelect.Click += Select_Click;

            // Position footer children after footer is sized
            footer.Layout += (s, e) =>
            {
                int cy = footer.Height / 2;
                _statusLabel.Location  = new Point(16, cy - _statusLabel.PreferredHeight / 2);
                newCaseLink.Location   = new Point(_statusLabel.Right + 12, cy - newCaseLink.PreferredHeight / 2);
                btnCancel.Location     = new Point(footer.Width - btnCancel.Width - btnSelect.Width - 12, cy - btnCancel.Height / 2);
                btnSelect.Location     = new Point(footer.Width - btnSelect.Width - 4, cy - btnSelect.Height / 2);
            };

            footer.Controls.AddRange(new Control[] { _statusLabel, newCaseLink, btnCancel, btnSelect });

            // ── Body panel (fills remaining space) ────────────────────────
            var body = new Panel
            {
                Dock      = DockStyle.Fill,
                BackColor = UITheme.Background,
                Padding   = new Padding(16, 12, 16, 12)
            };

            // Search box (plain TextBox — most reliable in WinForms)
            _searchBox = new TextBox
            {
                Dock        = DockStyle.Top,
                Height      = 32,
                Font        = UITheme.DefaultFont,
                BorderStyle = BorderStyle.FixedSingle,
                ForeColor   = UITheme.TextPrimary,
                BackColor   = Color.White
            };
            _searchBox.TextChanged += (s, e) => { _debounceTimer.Stop(); _debounceTimer.Start(); };

            // Spacer between search and list
            var spacer = new Panel { Dock = DockStyle.Top, Height = 8, BackColor = UITheme.Background };

            // List box (plain ListBox with owner-draw via DrawMode)
            _listBox = new ListBox
            {
                Dock          = DockStyle.Fill,
                Font          = UITheme.DefaultFont,
                BorderStyle   = BorderStyle.FixedSingle,
                BackColor     = Color.White,
                ForeColor     = UITheme.TextPrimary,
                ItemHeight    = 38,
                DrawMode      = DrawMode.OwnerDrawFixed,
                SelectionMode = SelectionMode.One,
                IntegralHeight = false
            };
            _listBox.DrawItem    += ListBox_DrawItem;
            _listBox.DoubleClick += Select_Click;

            // Add in reverse order for DockStyle.Fill + DockStyle.Top to work correctly
            body.Controls.Add(_listBox);   // Fill — added first so it fills remaining space
            body.Controls.Add(spacer);     // Top
            body.Controls.Add(_searchBox); // Top

            // ── Assemble form (order matters for Dock) ────────────────────
            // Bottom and Top docks must be added before Fill
            Controls.Add(body);     // Fill
            Controls.Add(footer);   // Bottom
            Controls.Add(titleBar); // Top

            AcceptButton = btnSelect;
            CancelButton = btnCancel;

            _debounceTimer = new System.Windows.Forms.Timer { Interval = 300 };
            _debounceTimer.Tick += DebounceTimer_Tick;

            Shown += async (s, e) => await SearchAsync(string.Empty);
        }

        // ── List drawing ──────────────────────────────────────────────────────

        private void ListBox_DrawItem(object sender, DrawItemEventArgs e)
        {
            if (e.Index < 0 || e.Index >= _listBox.Items.Count) return;

            var g = e.Graphics;
            bool selected = (e.State & DrawItemState.Selected) != 0;

            // Background
            Color bg = selected
                ? UITheme.AccentLight
                : (e.Index % 2 == 0 ? Color.White : Color.FromArgb(249, 250, 252));
            g.FillRectangle(new SolidBrush(bg), e.Bounds);

            // Accent left bar when selected
            if (selected)
                g.FillRectangle(new SolidBrush(UITheme.Accent),
                    e.Bounds.X, e.Bounds.Y, 3, e.Bounds.Height);

            // Text
            var text = _listBox.Items[e.Index]?.ToString() ?? string.Empty;
            var tf   = new RectangleF(e.Bounds.X + 14, e.Bounds.Y + 2,
                                      e.Bounds.Width - 18, e.Bounds.Height - 4);
            g.DrawString(text, UITheme.DefaultFont,
                new SolidBrush(selected ? UITheme.TextPrimary : UITheme.TextSecondary), tf,
                new StringFormat { LineAlignment = StringAlignment.Center, Trimming = StringTrimming.EllipsisCharacter });

            // Bottom separator
            if (!selected)
                g.DrawLine(new Pen(UITheme.Border),
                    e.Bounds.X + 10, e.Bounds.Bottom - 1,
                    e.Bounds.Right - 10, e.Bounds.Bottom - 1);
        }

        // ── Search ────────────────────────────────────────────────────────────

        private async void DebounceTimer_Tick(object sender, EventArgs e)
        {
            _debounceTimer.Stop();
            await SearchAsync(_searchBox.Text);
        }

        private async Task SearchAsync(string query)
        {
            try
            {
                _statusLabel.Text = "Suche …";
                _listBox.DataSource = null;

                var cases = await _api.SearchCasesAsync(query ?? string.Empty, CancellationToken.None)
                                      .ConfigureAwait(true);

                _listBox.DataSource = cases;
                _statusLabel.Text   = cases.Count + " Treffer";
            }
            catch (ProofioUnauthorizedException) { throw; }
            catch (Exception ex)
            {
                Logger.Error("Akten-Suche fehlgeschlagen.", ex);
                _statusLabel.Text = "Suche fehlgeschlagen.";
                MessageBox.Show("Die Akten-Suche ist fehlgeschlagen: " + ex.Message,
                    "Proofio", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // ── Actions ───────────────────────────────────────────────────────────

        private async void NewCaseLink_Click(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
                using (var dialog = new NewCaseDialog())
                {
                    if (dialog.ShowDialog(this) != DialogResult.OK) return;

                    var created = await _api.CreateCaseAsync(dialog.Request, CancellationToken.None)
                                            .ConfigureAwait(true);
                    SelectedCaseId = created.Id;
                    DialogResult   = DialogResult.OK;
                    Close();
                }
            }
            catch (Exception ex)
            {
                Logger.Error("Fall anlegen fehlgeschlagen.", ex);
                MessageBox.Show("Der Fall konnte nicht angelegt werden: " + ex.Message,
                    "Proofio", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Select_Click(object sender, EventArgs e)
        {
            var selected = _listBox.SelectedItem as CaseDto;
            if (selected == null || selected.Id == Guid.Empty)
            {
                MessageBox.Show("Bitte wählen Sie eine Akte aus.",
                    "Proofio", MessageBoxButtons.OK, MessageBoxIcon.Information);
                DialogResult = DialogResult.None;
                return;
            }
            SelectedCaseId = selected.Id;
            DialogResult   = DialogResult.OK;
            Close();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing) _debounceTimer?.Dispose();
            base.Dispose(disposing);
        }
    }
}
