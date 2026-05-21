using Microsoft.Office.Core;
using ProofioAddIn.Services;
using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;

namespace ProofioAddIn.Ribbons
{
    [ComVisible(true)]
    public sealed class ProofioRibbonManager : IRibbonExtensibility
    {
        private readonly ExplorerRibbon _explorer = new ExplorerRibbon();
        private readonly MailReadRibbon _mailRead = new MailReadRibbon();
        private readonly MailComposeRibbon _mailCompose = new MailComposeRibbon();
        private readonly AppointmentRibbon _appointment = new AppointmentRibbon();

        public string GetCustomUI(string ribbonID)
        {
            Logger.Info("Ribbon angefordert: " + ribbonID);

            switch (ribbonID)
            {
                case "Microsoft.Outlook.Explorer":
                    return LoadXml("ExplorerRibbon.xml");

                case "Microsoft.Outlook.Mail.Read":
                    return LoadXml("MailReadRibbon.xml");

                case "Microsoft.Outlook.Mail.Compose":
                    return LoadXml("MailComposeRibbon.xml");

                case "Microsoft.Outlook.Appointment":
                    return LoadXml("AppointmentRibbon.xml");

                default:
                    return null;
            }
        }

        public void OnLoad(IRibbonUI ribbon)
        {
            _explorer.Ribbon = ribbon;
            _mailRead.Ribbon = ribbon;
            _mailCompose.Ribbon = ribbon;
            _appointment.Ribbon = ribbon;
        }

        public void OnSettings(IRibbonControl control) { _explorer.OnSettings(control); }
        public void OnFileSelectedMails(IRibbonControl control) { _explorer.OnFileSelectedMails(control); }
        public void OnFileOpenMail(IRibbonControl control) { _mailRead.OnFileOpenMail(control); }
        public void OnSendAndFile(IRibbonControl control) { _mailCompose.OnSendAndFile(control); }
        public bool GetFileAfterSendPressed(IRibbonControl control) { return _mailCompose.GetFileAfterSendPressed(control); }
        public void OnFileAfterSendChanged(IRibbonControl control, bool pressed) { _mailCompose.OnFileAfterSendChanged(control, pressed); }
        public void OnFileAppointment(IRibbonControl control) { _appointment.OnFileAppointment(control); }

        private static string LoadXml(string name)
        {
            var asm = Assembly.GetExecutingAssembly();
            var expectedSuffix = ".Ribbons." + name;

            foreach (var resourceName in asm.GetManifestResourceNames())
            {
                if (resourceName.EndsWith(expectedSuffix, StringComparison.OrdinalIgnoreCase))
                {
                    using (var stream = asm.GetManifestResourceStream(resourceName))
                    using (var reader = new StreamReader(stream, Encoding.UTF8))
                    {
                        return reader.ReadToEnd();
                    }
                }
            }

            throw new InvalidOperationException("Ribbon-XML nicht gefunden: " + expectedSuffix);
        }
    }
}
