using ProofioAddIn.Api;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace ProofioAddIn.Services
{
    public sealed class AppointmentExtractor
    {
        private const string PrSmtpAddress = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E";

        public AppointmentPayload Extract(Outlook.AppointmentItem appointment, Guid caseId)
        {
            if (appointment == null)
            {
                throw new ArgumentNullException(nameof(appointment));
            }

            var subject = SafeString(appointment.Subject);
            if (string.IsNullOrWhiteSpace(subject))
            {
                subject = "(ohne Betreff)";
            }

            var start = GetStart(appointment);
            var end = GetEnd(appointment, start);

            var payload = new AppointmentPayload
            {
                CaseId = caseId,
                Subject = subject,
                Location = SafeString(appointment.Location),
                StartsAt = start.ToUniversalTime().ToString("o"),
                EndsAt = end.ToUniversalTime().ToString("o"),
                Notes = GetNotes(appointment),
                Organizer = GetOrganizer(appointment),
                Attendees = ExtractAttendees(appointment)
            };

            if (payload.Attendees == null)
            {
                payload.Attendees = new List<AppointmentAttendeePayload>();
            }

            Logger.Info(
                "Termin extrahiert: Subject='" + payload.Subject +
                "', Organizer='" + (payload.Organizer == null ? "" : payload.Organizer.Email) +
                "', Attendees=" + payload.Attendees.Count +
                ", NotesLength=" + (payload.Notes == null ? 0 : payload.Notes.Length));

            return payload;
        }

        private static DateTime GetStart(Outlook.AppointmentItem appointment)
        {
            try
            {
                var start = appointment.Start;
                if (start != DateTime.MinValue) return start;
            }
            catch { }

            return DateTime.Now;
        }

        private static DateTime GetEnd(Outlook.AppointmentItem appointment, DateTime start)
        {
            try
            {
                var end = appointment.End;
                if (end != DateTime.MinValue && end > start) return end;
            }
            catch { }

            return start.AddHours(1);
        }

        private static string GetNotes(Outlook.AppointmentItem appointment)
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(appointment.Body)) return appointment.Body;
            }
            catch { }

            return string.Empty;
        }

        private static AppointmentOrganizerPayload GetOrganizer(Outlook.AppointmentItem appointment)
        {
            var organizer = new AppointmentOrganizerPayload();

            try
            {
                if (!string.IsNullOrWhiteSpace(appointment.Organizer))
                {
                    organizer.Name = appointment.Organizer;
                }
            }
            catch { }

            try
            {
                var session = appointment.Application.Session;
                var currentUser = session.CurrentUser;

                if (currentUser != null)
                {
                    if (string.IsNullOrWhiteSpace(organizer.Name))
                    {
                        organizer.Name = currentUser.Name;
                    }

                    organizer.Email = ResolveAddressEntry(currentUser.AddressEntry);
                }
            }
            catch { }

            if (string.IsNullOrWhiteSpace(organizer.Name) &&
                string.IsNullOrWhiteSpace(organizer.Email))
            {
                return null;
            }

            return organizer;
        }

        private static List<AppointmentAttendeePayload> ExtractAttendees(Outlook.AppointmentItem appointment)
        {
            var result = new List<AppointmentAttendeePayload>();
            Outlook.Recipients recipients = null;

            try
            {
                recipients = appointment.Recipients;
                if (recipients == null || recipients.Count == 0) return result;

                for (var i = 1; i <= recipients.Count; i++)
                {
                    Outlook.Recipient recipient = null;

                    try
                    {
                        recipient = recipients[i];
                        if (recipient == null) continue;

                        var attendee = new AppointmentAttendeePayload
                        {
                            Name = SafeString(recipient.Name),
                            Email = ResolveRecipientEmail(recipient),
                            Type = ResolveRecipientType(recipient),
                            ResponseStatus = ResolveResponseStatus(recipient)
                        };

                        if (!string.IsNullOrWhiteSpace(attendee.Name) ||
                            !string.IsNullOrWhiteSpace(attendee.Email))
                        {
                            result.Add(attendee);
                        }
                    }
                    catch (Exception ex)
                    {
                        Logger.Error("Termin-Teilnehmer konnte nicht extrahiert werden.", ex);
                    }
                    finally
                    {
                        if (recipient != null) Marshal.ReleaseComObject(recipient);
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Error("Termin-Teilnehmer konnten nicht gelesen werden.", ex);
            }
            finally
            {
                if (recipients != null) Marshal.ReleaseComObject(recipients);
            }

            return result;
        }

        private static string ResolveRecipientEmail(Outlook.Recipient recipient)
        {
            try
            {
                if (!recipient.Resolved) recipient.Resolve();
            }
            catch { }

            try
            {
                var entry = recipient.AddressEntry;
                var smtp = ResolveAddressEntry(entry);
                if (!string.IsNullOrWhiteSpace(smtp)) return smtp;
            }
            catch { }

            try
            {
                if (!string.IsNullOrWhiteSpace(recipient.Address) &&
                    !LooksLikeExchangeLegacyDn(recipient.Address))
                {
                    return recipient.Address;
                }
            }
            catch { }

            return string.Empty;
        }

        private static string ResolveAddressEntry(Outlook.AddressEntry entry)
        {
            if (entry == null) return string.Empty;

            try
            {
                if (string.Equals(entry.Type, "SMTP", StringComparison.OrdinalIgnoreCase) &&
                    !string.IsNullOrWhiteSpace(entry.Address))
                {
                    return entry.Address;
                }
            }
            catch { }

            try
            {
                var exchangeUser = entry.GetExchangeUser();
                if (exchangeUser != null && !string.IsNullOrWhiteSpace(exchangeUser.PrimarySmtpAddress))
                {
                    return exchangeUser.PrimarySmtpAddress;
                }
            }
            catch { }

            try
            {
                var exchangeDistributionList = entry.GetExchangeDistributionList();
                if (exchangeDistributionList != null &&
                    !string.IsNullOrWhiteSpace(exchangeDistributionList.PrimarySmtpAddress))
                {
                    return exchangeDistributionList.PrimarySmtpAddress;
                }
            }
            catch { }

            try
            {
                var smtp = entry.PropertyAccessor.GetProperty(PrSmtpAddress) as string;
                if (!string.IsNullOrWhiteSpace(smtp)) return smtp;
            }
            catch { }

            try
            {
                if (!string.IsNullOrWhiteSpace(entry.Address) &&
                    !LooksLikeExchangeLegacyDn(entry.Address))
                {
                    return entry.Address;
                }
            }
            catch { }

            return string.Empty;
        }

        private static string ResolveRecipientType(Outlook.Recipient recipient)
        {
            try
            {
                switch (recipient.Type)
                {
                    case 1: return "required";
                    case 2: return "optional";
                    case 3: return "resource";
                    default: return "unknown";
                }
            }
            catch
            {
                return "unknown";
            }
        }

        private static string ResolveResponseStatus(Outlook.Recipient recipient)
        {
            try
            {
                switch (recipient.MeetingResponseStatus)
                {
                    case Outlook.OlResponseStatus.olResponseAccepted: return "accepted";
                    case Outlook.OlResponseStatus.olResponseDeclined: return "declined";
                    case Outlook.OlResponseStatus.olResponseTentative: return "tentative";
                    case Outlook.OlResponseStatus.olResponseNotResponded: return "notResponded";
                    case Outlook.OlResponseStatus.olResponseOrganized: return "organized";
                    default: return "unknown";
                }
            }
            catch
            {
                return "unknown";
            }
        }

        private static bool LooksLikeExchangeLegacyDn(string value)
        {
            if (string.IsNullOrWhiteSpace(value)) return false;

            return value.StartsWith("/O=", StringComparison.OrdinalIgnoreCase) ||
                   value.IndexOf("/OU=", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   value.IndexOf("/CN=RECIPIENTS/", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static string SafeString(string value)
        {
            return value ?? string.Empty;
        }
    }
}
