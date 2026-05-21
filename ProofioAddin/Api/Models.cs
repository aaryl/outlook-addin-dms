using System;
using System.Collections.Generic;
using System.Runtime.Serialization;

namespace ProofioAddIn.Api
{
    // ── Case search ──────────────────────────────────────────────────────────

    [DataContract]
    public sealed class CaseSearchResponse
    {
        [DataMember(Name = "cases")]
        public List<CaseDto> Cases { get; set; }
    }

    // ── Create case ──────────────────────────────────────────────────────────

    /// <summary>
    /// POST /api/public/v1/outlook/case
    /// Only <see cref="Title"/> is required; all other fields are optional.
    /// </summary>
    [DataContract]
    public sealed class CreateCaseRequest
    {
        /// <summary>1–300 characters, required.</summary>
        [DataMember(Name = "title")]
        public string Title { get; set; }

        /// <summary>gerichtsgutachten | privatgutachten | schiedsgutachten | parteigutachten</summary>
        [DataMember(Name = "art", EmitDefaultValue = false)]
        public string Art { get; set; }

        [DataMember(Name = "auftraggeber_email", EmitDefaultValue = false)]
        public string AuftraggeberEmail { get; set; }

        [DataMember(Name = "auftraggeber_vorname", EmitDefaultValue = false)]
        public string AuftraggeberVorname { get; set; }

        [DataMember(Name = "auftraggeber_nachname", EmitDefaultValue = false)]
        public string AuftraggeberNachname { get; set; }

        [DataMember(Name = "auftraggeber_firma", EmitDefaultValue = false)]
        public string AuftraggeberFirma { get; set; }
    }

    /// <summary>
    /// Response body: { "ok": true, "case": { ... } }
    /// </summary>
    [DataContract]
    public sealed class CreateCaseResponse
    {
        [DataMember(Name = "ok")]
        public bool Ok { get; set; }

        [DataMember(Name = "case")]
        public CaseDto Case { get; set; }
    }

    // ── Shared DTO ───────────────────────────────────────────────────────────

    [DataContract]
    public sealed class CaseDto
    {
        [DataMember(Name = "id")]
        public Guid Id { get; set; }

        /// <summary>Automatically assigned by the server, e.g. "P-2026-0042".</summary>
        [DataMember(Name = "aktenzeichen")]
        public string Aktenzeichen { get; set; }

        [DataMember(Name = "title")]
        public string Title { get; set; }

        [DataMember(Name = "art")]
        public string Art { get; set; }

        [DataMember(Name = "status")]
        public string Status { get; set; }

        [DataMember(Name = "updated_at")]
        public string UpdatedAt { get; set; }

        public override string ToString()
        {
            var az    = string.IsNullOrWhiteSpace(Aktenzeichen) ? "(ohne AZ)"    : Aktenzeichen;
            var title = string.IsNullOrWhiteSpace(Title)        ? "(ohne Titel)" : Title;
            return az + " — " + title;
        }
    }

    // ── Email / Appointment payloads (unchanged) ─────────────────────────────

    [DataContract]
    public sealed class EmailPayload
    {
        public EmailPayload()
        {
            ToAddrs     = new List<string>();
            CcAddrs     = new List<string>();
            Attachments = new List<AttachmentPayload>();
        }

        [DataMember(Name = "caseId")]
        public Guid CaseId { get; set; }

        [DataMember(Name = "subject")]
        public string Subject { get; set; }

        [DataMember(Name = "fromAddr")]
        public string FromAddr { get; set; }

        [DataMember(Name = "toAddrs")]
        public List<string> ToAddrs { get; set; }

        [DataMember(Name = "ccAddrs")]
        public List<string> CcAddrs { get; set; }

        [DataMember(Name = "messageDate")]
        public string MessageDate { get; set; }

        [DataMember(Name = "direction", EmitDefaultValue = false)]
        public string Direction { get; set; }

        [DataMember(Name = "bodyText", EmitDefaultValue = false)]
        public string BodyText { get; set; }

        [DataMember(Name = "bodyHtml")]
        public string BodyHtml { get; set; }

        [DataMember(Name = "attachments")]
        public List<AttachmentPayload> Attachments { get; set; }
    }

    [DataContract]
    public sealed class AttachmentPayload
    {
        [DataMember(Name = "name")]
        public string Name { get; set; }

        [DataMember(Name = "contentType")]
        public string ContentType { get; set; }

        [DataMember(Name = "contentBase64")]
        public string ContentBase64 { get; set; }
    }

    [DataContract]
    public sealed class AppointmentPayload
    {
        public AppointmentPayload()
        {
            Attendees = new List<AppointmentAttendeePayload>();
        }

        [DataMember(Name = "caseId")]
        public Guid CaseId { get; set; }

        [DataMember(Name = "title")]
        public string Subject { get; set; }

        [DataMember(Name = "location")]
        public string Location { get; set; }

        [DataMember(Name = "startsAt")]
        public string StartsAt { get; set; }

        [DataMember(Name = "endsAt")]
        public string EndsAt { get; set; }

        [DataMember(Name = "notes")]
        public string Notes { get; set; }

        [DataMember(Name = "organizer", EmitDefaultValue = false)]
        public AppointmentOrganizerPayload Organizer { get; set; }

        [DataMember(Name = "attendees")]
        public List<AppointmentAttendeePayload> Attendees { get; set; }
    }

    [DataContract]
    public sealed class AppointmentOrganizerPayload
    {
        [DataMember(Name = "name")]
        public string Name { get; set; }

        [DataMember(Name = "email")]
        public string Email { get; set; }
    }

    [DataContract]
    public sealed class AppointmentAttendeePayload
    {
        [DataMember(Name = "name")]
        public string Name { get; set; }

        [DataMember(Name = "email")]
        public string Email { get; set; }

        [DataMember(Name = "type")]
        public string Type { get; set; }

        [DataMember(Name = "responseStatus")]
        public string ResponseStatus { get; set; }
    }
}
