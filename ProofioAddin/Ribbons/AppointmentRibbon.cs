using Microsoft.Office.Core;
using ProofioAddIn.Api;
using ProofioAddIn.UI;
using System;
using System.Threading;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace ProofioAddIn.Ribbons
{
    public sealed class AppointmentRibbon
    {
        public IRibbonUI Ribbon { get; set; }

        public async void OnFileAppointment(IRibbonControl control)
        {
            try
            {
                var inspector = Globals.ThisAddIn.Application.ActiveInspector();
                var appointment = inspector != null ? inspector.CurrentItem as Outlook.AppointmentItem : null;

                if (appointment == null)
                {
                    MessageBox.Show("Das geöffnete Element ist kein Termin.", "Proofio", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                using (var picker = new CasePickerDialog(Globals.ThisAddIn.ApiClient))
                {
                    if (picker.ShowDialog() != DialogResult.OK || picker.SelectedCaseId == Guid.Empty) return;

                    await Globals.ThisAddIn.FileAppointmentAsync(appointment, picker.SelectedCaseId, CancellationToken.None).ConfigureAwait(false);

                    MessageBox.Show("Der Termin wurde in Proofio übernommen.", "Proofio", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (ProofioUnauthorizedException ex)
            {
                MessageBox.Show(ex.Message, "Proofio", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                Globals.ThisAddIn.ShowTokenDialog();
            }
            catch (Exception ex)
            {
                ProofioAddIn.Services.Logger.Error("Termin ablegen fehlgeschlagen.", ex);
                MessageBox.Show("Der Termin konnte nicht übernommen werden: " + ex.Message, "Proofio", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
