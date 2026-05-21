using Microsoft.Office.Core;
using ProofioAddIn.Api;
using ProofioAddIn.UI;
using System;
using System.Threading;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace ProofioAddIn.Ribbons
{
    public sealed class MailReadRibbon
    {
        public IRibbonUI Ribbon { get; set; }

        public async void OnFileOpenMail(IRibbonControl control)
        {
            try
            {
                var inspector = Globals.ThisAddIn.Application.ActiveInspector();
                var mail = inspector != null ? inspector.CurrentItem as Outlook.MailItem : null;

                if (mail == null)
                {
                    MessageBox.Show("Das geöffnete Element ist keine Mail.", "Proofio", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                using (var picker = new CasePickerDialog(Globals.ThisAddIn.ApiClient))
                {
                    if (picker.ShowDialog() != DialogResult.OK || picker.SelectedCaseId == Guid.Empty) return;

                    await Globals.ThisAddIn.FileMailAsync(mail, picker.SelectedCaseId, CancellationToken.None).ConfigureAwait(false);

                    MessageBox.Show("Die Mail wurde in Proofio abgelegt.", "Proofio", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (ProofioUnauthorizedException ex)
            {
                MessageBox.Show(ex.Message, "Proofio", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                Globals.ThisAddIn.ShowTokenDialog();
            }
            catch (Exception ex)
            {
                ProofioAddIn.Services.Logger.Error("Geöffnete Mail ablegen fehlgeschlagen.", ex);
                MessageBox.Show("Die Mail konnte nicht abgelegt werden: " + ex.Message, "Proofio", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
