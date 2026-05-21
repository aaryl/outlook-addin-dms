using ProofioAddIn.Api;
using System;
using System.Collections.Generic;
using System.Threading;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace ProofioAddIn.Services
{
    public sealed class PendingSendTracker : IDisposable
    {
        public const string FileAfterSendPropertyName = "Proofio.FileAfterSend";
        public const string PendingCaseIdPropertyName = "Proofio.PendingCaseId";

        private readonly ProofioApiClient _api;
        private readonly MailExtractor _extractor;
        private readonly Dictionary<string, EmailPayload> _pending =
            new Dictionary<string, EmailPayload>(StringComparer.OrdinalIgnoreCase);

        private Outlook.Items _sentItems;

        public PendingSendTracker(ProofioApiClient api, MailExtractor extractor)
        {
            _api = api ?? throw new ArgumentNullException(nameof(api));
            _extractor = extractor ?? throw new ArgumentNullException(nameof(extractor));
        }

        public void Start(Outlook.Application application)
        {
            var ns = application.Session;
            var sentFolder = ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderSentMail);

            _sentItems = sentFolder.Items;
            _sentItems.ItemAdd += SentItems_ItemAdd;

            Logger.Info("PendingSendTracker gestartet.");
        }

        public void TrackBeforeSend(Outlook.MailItem mail, Guid caseId)
        {
            var payload = _extractor.Extract(mail, caseId, "ausgehend");
            var key = CreateKey(mail);

            _pending[key] = payload;

            SetUserProperty(mail, PendingCaseIdPropertyName, caseId.ToString());

            Logger.Info("Mail für Ablage nach Senden vorgemerkt: " + key);
        }

        private async void SentItems_ItemAdd(object item)
        {
            try
            {
                var mail = item as Outlook.MailItem;
                if (mail == null) return;

                var key = CreateKey(mail);

                EmailPayload payload;
                if (!_pending.TryGetValue(key, out payload))
                {
                    var caseIdText = GetUserProperty(mail, PendingCaseIdPropertyName) as string;

                    Guid caseId;
                    if (Guid.TryParse(caseIdText, out caseId))
                    {
                        payload = _extractor.Extract(mail, caseId, "ausgehend");
                    }
                }

                if (payload == null) return;

                await _api.FileEmailAsync(payload, CancellationToken.None).ConfigureAwait(false);

                _pending.Remove(key);

                Logger.Info("Gesendete Mail erfolgreich in Proofio abgelegt: " + key);
            }
            catch (ProofioUnauthorizedException ex)
            {
                Logger.Error("Ablage nach Senden nicht autorisiert.", ex);
                MessageBox.Show(ex.Message, "Proofio", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            catch (Exception ex)
            {
                Logger.Error("Ablage nach Senden fehlgeschlagen.", ex);
                MessageBox.Show("Die gesendete Mail konnte nicht in Proofio abgelegt werden: " + ex.Message, "Proofio", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public static bool GetFileAfterSend(Outlook.MailItem mail)
        {
            var value = GetUserProperty(mail, FileAfterSendPropertyName);

            if (value is bool b) return b;

            bool parsed;
            return bool.TryParse(Convert.ToString(value), out parsed) && parsed;
        }

        public static void SetFileAfterSend(Outlook.MailItem mail, bool value)
        {
            SetUserProperty(mail, FileAfterSendPropertyName, value);

            try { mail.Save(); } catch { }
        }

        private static object GetUserProperty(Outlook.MailItem mail, string name)
        {
            try
            {
                var prop = mail.UserProperties.Find(name, true);
                return prop == null ? null : prop.Value;
            }
            catch
            {
                return null;
            }
        }

        private static void SetUserProperty(Outlook.MailItem mail, string name, object value)
        {
            var prop = mail.UserProperties.Find(name, true);

            if (prop == null)
            {
                var type = value is bool
                    ? Outlook.OlUserPropertyType.olYesNo
                    : Outlook.OlUserPropertyType.olText;

                prop = mail.UserProperties.Add(name, type, false);
            }

            prop.Value = value;
        }

        private static string CreateKey(Outlook.MailItem mail)
        {
            var subject = mail.Subject ?? string.Empty;
            var to = mail.To ?? string.Empty;
            var cc = mail.CC ?? string.Empty;
            var stamp = DateTimeOffset.Now.ToString("yyyyMMddHHmm");

            try
            {
                if (!string.IsNullOrWhiteSpace(mail.EntryID)) stamp = mail.EntryID;
            }
            catch { }

            return subject + "|" + to + "|" + cc + "|" + stamp;
        }

        public void Dispose()
        {
            if (_sentItems != null)
            {
                _sentItems.ItemAdd -= SentItems_ItemAdd;
                _sentItems = null;
            }
        }
    }
}
