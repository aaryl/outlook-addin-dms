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
    public sealed class MailComposeRibbon
    {
        public IRibbonUI Ribbon { get; set; }

        public async void OnSendAndFile(IRibbonControl control)
        {
            try
            {
                var mail = GetCurrentMail();
                if (mail == null)
                {
                    MessageBox.Show(
                        "Das geöffnete Element ist keine Mail im Verfassen-Modus.",
                        "Proofio",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);

                    return;
                }

                using (var picker = new CasePickerDialog(Globals.ThisAddIn.ApiClient))
                {
                    if (picker.ShowDialog() != DialogResult.OK || picker.SelectedCaseId == Guid.Empty) return;

                    PendingSendTracker.SetFileAfterSend(mail, false);

                    if (Ribbon != null) Ribbon.InvalidateControl("proofioFileAfterSend");

                    var payload = Globals.ThisAddIn.MailExtractor.Extract(
                        mail,
                        picker.SelectedCaseId,
                        "ausgehend");

                    mail.Send();

                    await Globals.ThisAddIn.ApiClient.FileEmailAsync(
                        payload,
                        CancellationToken.None).ConfigureAwait(false);

                    Services.MailCategoryService.Apply(mail); // grünes Badge

                    MessageBox.Show("Die Mail wurde gesendet und in Proofio abgelegt.", "Proofio", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (ProofioUnauthorizedException ex)
            {
                MessageBox.Show(ex.Message, "Proofio", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                Globals.ThisAddIn.ShowTokenDialog();
            }
            catch (Exception ex)
            {
                Services.Logger.Error("Senden und ablegen fehlgeschlagen.", ex);
                MessageBox.Show("Senden und Ablegen ist fehlgeschlagen: " + ex.Message, "Proofio", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public bool GetFileAfterSendPressed(IRibbonControl control)
        {
            var mail = GetCurrentMail();
            return mail != null && PendingSendTracker.GetFileAfterSend(mail);
        }

        public void OnFileAfterSendChanged(IRibbonControl control, bool pressed)
        {
            try
            {
                var mail = GetCurrentMail();
                if (mail == null) return;

                PendingSendTracker.SetFileAfterSend(mail, pressed);

                if (Ribbon != null) Ribbon.InvalidateControl("proofioFileAfterSend");
            }
            catch (Exception ex)
            {
                Services.Logger.Error("Toggle Nach Senden ablegen fehlgeschlagen.", ex);
                MessageBox.Show("Die Einstellung konnte nicht in der Mail gespeichert werden: " + ex.Message, "Proofio", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private static Outlook.MailItem GetCurrentMail()
        {
            var inspector = Globals.ThisAddIn.Application.ActiveInspector();
            return inspector != null ? inspector.CurrentItem as Outlook.MailItem : null;
        }
    }
}
