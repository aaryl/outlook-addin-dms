using Microsoft.Office.Core;
using ProofioAddIn.Api;
using ProofioAddIn.Ribbons;
using ProofioAddIn.Services;
using ProofioAddIn.UI;
using System;
using System.Net;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace ProofioAddIn
{
    public partial class ThisAddIn
    {
        public TokenStore TokenStore { get; private set; }
        public ProofioApiClient ApiClient { get; private set; }
        public MailExtractor MailExtractor { get; private set; }
        public AppointmentExtractor AppointmentExtractor { get; private set; }
        public PendingSendTracker PendingSendTracker { get; private set; }

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            try
            {
                ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;

                TokenStore = new TokenStore();
                ApiClient = new ProofioApiClient(TokenStore);
                MailExtractor = new MailExtractor();
                AppointmentExtractor = new AppointmentExtractor();
                PendingSendTracker = new PendingSendTracker(ApiClient, MailExtractor);

                PendingSendTracker.Start(this.Application);
                this.Application.ItemSend += Application_ItemSend;

                Logger.Info("ProofioAddIn gestartet.");
            }
            catch (Exception ex)
            {
                Logger.Error("ProofioAddIn konnte nicht gestartet werden.", ex);

                MessageBox.Show(
                    "Proofio konnte nicht gestartet werden: " + ex.Message,
                    "Proofio",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            try
            {
                this.Application.ItemSend -= Application_ItemSend;
                PendingSendTracker?.Dispose();
                Logger.Info("ProofioAddIn beendet.");
            }
            catch
            {
            }
        }

        protected override IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new ProofioRibbonManager();
        }

        private void Application_ItemSend(object item, ref bool cancel)
        {
            try
            {
                var mail = item as Outlook.MailItem;
                if (mail == null)
                {
                    return;
                }

                var shouldFile = PendingSendTracker.GetFileAfterSend(mail);
                var mode = TokenStore.GetSendFilingMode();

                if (!shouldFile)
                {
                    if (string.Equals(mode, TokenStore.SendFilingModeNever, StringComparison.OrdinalIgnoreCase))
                    {
                        return;
                    }

                    if (string.Equals(mode, TokenStore.SendFilingModeAsk, StringComparison.OrdinalIgnoreCase))
                    {
                        var answer = MessageBox.Show(
                            "Möchten Sie diese E-Mail in Proofio ablegen?",
                            "Proofio",
                            MessageBoxButtons.YesNo,
                            MessageBoxIcon.Question);

                        if (answer != DialogResult.Yes)
                        {
                            return;
                        }

                        shouldFile = true;
                    }
                    else if (string.Equals(mode, TokenStore.SendFilingModeAlways, StringComparison.OrdinalIgnoreCase))
                    {
                        shouldFile = true;
                    }
                }

                if (!shouldFile)
                {
                    return;
                }

                using (var picker = new CasePickerDialog(ApiClient))
                {
                    if (picker.ShowDialog() != DialogResult.OK || picker.SelectedCaseId == Guid.Empty)
                    {
                        cancel = true;

                        MessageBox.Show(
                            "Senden wurde abgebrochen, weil keine Akte ausgewählt wurde.",
                            "Proofio",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information);

                        return;
                    }

                    /*
                     * ItemSend ist ein echter Send-Workflow.
                     * PendingSendTracker erzwingt beim Payload direction = "ausgehend".
                     */
                    PendingSendTracker.TrackBeforeSend(mail, picker.SelectedCaseId);
                }
            }
            catch (ProofioUnauthorizedException ex)
            {
                cancel = true;

                Logger.Error("ItemSend: Token ungültig.", ex);

                MessageBox.Show(
                    ex.Message,
                    "Proofio",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);

                ShowTokenDialog();
            }
            catch (Exception ex)
            {
                cancel = true;

                Logger.Error("ItemSend fehlgeschlagen.", ex);

                MessageBox.Show(
                    "Die Mail konnte nicht für Proofio vorbereitet werden: " + ex.Message,
                    "Proofio",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        public async Task FileMailAsync(Outlook.MailItem mail, Guid caseId, CancellationToken cancellationToken)
        {
            /*
             * Normales manuelles Ablegen aus dem Postfach/Explorer/Lesefenster.
             * Deshalb ohne forcedDirection → "eingehend".
             */
            var payload = MailExtractor.Extract(mail, caseId);
            await ApiClient.FileEmailAsync(payload, cancellationToken).ConfigureAwait(false);
            MailCategoryService.Apply(mail); // grünes Badge im Posteingang
        }

        public async Task FileOutgoingMailAsync(Outlook.MailItem mail, Guid caseId, CancellationToken cancellationToken)
        {
            /*
             * Nur für echte Sende-Workflows.
             */
            var payload = MailExtractor.Extract(mail, caseId, "ausgehend");
            await ApiClient.FileEmailAsync(payload, cancellationToken).ConfigureAwait(false);
            MailCategoryService.Apply(mail); // grünes Badge im Posteingang
        }

        public async Task FileAppointmentAsync(Outlook.AppointmentItem appointment, Guid caseId, CancellationToken cancellationToken)
        {
            var payload = AppointmentExtractor.Extract(appointment, caseId);
            await ApiClient.FileAppointmentAsync(payload, cancellationToken).ConfigureAwait(false);
        }

        public void ShowTokenDialog()
        {
            using (var dialog = new TokenDialog(TokenStore))
            {
                dialog.ShowDialog();
            }
        }

        #region VSTO generated code

        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}