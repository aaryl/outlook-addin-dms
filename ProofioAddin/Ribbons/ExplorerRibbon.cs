using Microsoft.Office.Core;
using ProofioAddIn.Api;
using ProofioAddIn.Services;
using ProofioAddIn.UI;
using System;
using System.Threading;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace ProofioAddIn.Ribbons
{
    public sealed class ExplorerRibbon
    {
        public IRibbonUI Ribbon { get; set; }

        public void OnSettings(IRibbonControl control)
        {
            Globals.ThisAddIn.ShowTokenDialog();
        }

        public async void OnFileSelectedMails(IRibbonControl control)
        {
            try
            {
                var explorer = Globals.ThisAddIn.Application.ActiveExplorer();
                if (explorer == null || explorer.Selection == null || explorer.Selection.Count == 0)
                {
                    MessageBox.Show("Bitte markieren Sie zuerst mindestens eine Mail.", "Proofio", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                using (var picker = new CasePickerDialog(Globals.ThisAddIn.ApiClient))
                {
                    if (picker.ShowDialog() != DialogResult.OK || picker.SelectedCaseId == Guid.Empty) return;

                    var count = 0;
                    for (var i = 1; i <= explorer.Selection.Count; i++)
                    {
                        var mail = explorer.Selection[i] as Outlook.MailItem;
                        if (mail == null) continue;

                        await Globals.ThisAddIn.FileMailAsync(mail, picker.SelectedCaseId, CancellationToken.None).ConfigureAwait(false);
                        count++;
                    }

                    MessageBox.Show(count + " Mail(s) wurden in Proofio abgelegt.", "Proofio", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (ProofioUnauthorizedException ex)
            {
                MessageBox.Show(ex.Message, "Proofio", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                Globals.ThisAddIn.ShowTokenDialog();
            }
            catch (Exception ex)
            {
                Services.Logger.Error("Markierte Mails ablegen fehlgeschlagen.", ex);
                MessageBox.Show("Die markierten Mails konnten nicht abgelegt werden: " + ex.Message, "Proofio", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
