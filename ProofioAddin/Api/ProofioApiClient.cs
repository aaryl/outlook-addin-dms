using ProofioAddIn.Services;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace ProofioAddIn.Api
{
    public sealed class ProofioUnauthorizedException : Exception
    {
        public ProofioUnauthorizedException()
            : base("Token ungültig – bitte in den Einstellungen aktualisieren.")
        { }
    }

    public sealed class ProofioApiClient
    {
        private static readonly Lazy<HttpClient> SharedClient =
            new Lazy<HttpClient>(() => new HttpClient { Timeout = TimeSpan.FromSeconds(60) });

        private readonly TokenStore _tokenStore;

        public ProofioApiClient(TokenStore tokenStore)
        {
            _tokenStore = tokenStore ?? throw new ArgumentNullException(nameof(tokenStore));
        }

        // ── Public API ────────────────────────────────────────────────────────

        public async Task<List<CaseDto>> SearchCasesAsync(string query, CancellationToken ct)
        {
            var path = "/cases/search?q=" + Uri.EscapeDataString(query ?? string.Empty) + "&limit=20";
            var response = await SendAsync<CaseSearchResponse>(HttpMethod.Get, path, null, ct).ConfigureAwait(false);
            return response?.Cases ?? new List<CaseDto>();
        }

        /// <summary>
        /// Creates a new case via POST /api/public/v1/outlook/case.
        /// Only <paramref name="request"/>.Title is required; all other fields are optional.
        /// </summary>
        public async Task<CaseDto> CreateCaseAsync(CreateCaseRequest request, CancellationToken ct)
        {
            if (request == null) throw new ArgumentNullException(nameof(request));
            if (string.IsNullOrWhiteSpace(request.Title))
                throw new ArgumentException("Titel darf nicht leer sein.", nameof(request));

            // Build JSON manually to avoid DataContractJsonSerializer quirks
            var json = BuildCreateCaseJson(request);
            var response = await SendRawAsync<CreateCaseResponse>(
                HttpMethod.Post, "/outlook/case", json, ct).ConfigureAwait(false);

            if (response?.Case == null)
                throw new InvalidOperationException("Proofio API hat keinen Fall zurückgegeben.");

            return response.Case;
        }

        /// <summary>Builds a clean JSON string for CreateCaseRequest.</summary>
        private static string BuildCreateCaseJson(CreateCaseRequest r)
        {
            var sb = new System.Text.StringBuilder();
            sb.Append('{');
            sb.AppendFormat("\"title\":{0}", JsonString(r.Title));
            if (!string.IsNullOrWhiteSpace(r.Art))
                sb.AppendFormat(",\"art\":{0}", JsonString(r.Art));
            if (!string.IsNullOrWhiteSpace(r.AuftraggeberEmail))
                sb.AppendFormat(",\"auftraggeber_email\":{0}", JsonString(r.AuftraggeberEmail));
            if (!string.IsNullOrWhiteSpace(r.AuftraggeberVorname))
                sb.AppendFormat(",\"auftraggeber_vorname\":{0}", JsonString(r.AuftraggeberVorname));
            if (!string.IsNullOrWhiteSpace(r.AuftraggeberNachname))
                sb.AppendFormat(",\"auftraggeber_nachname\":{0}", JsonString(r.AuftraggeberNachname));
            if (!string.IsNullOrWhiteSpace(r.AuftraggeberFirma))
                sb.AppendFormat(",\"auftraggeber_firma\":{0}", JsonString(r.AuftraggeberFirma));
            sb.Append('}');
            return sb.ToString();
        }

        private static string JsonString(string s)
        {
            if (s == null) return "null";
            return "\"" + s.Replace("\\", "\\\\").Replace("\"", "\\\"").Replace("\n", "\\n").Replace("\r", "\\r").Replace("\t", "\\t") + "\"";
        }

        public async Task FileEmailAsync(EmailPayload payload, CancellationToken ct)
        {
            await SendAsync<object>(HttpMethod.Post, "/outlook/email", payload, ct).ConfigureAwait(false);
        }

        public async Task FileAppointmentAsync(AppointmentPayload payload, CancellationToken ct)
        {
            await SendAsync<object>(HttpMethod.Post, "/outlook/appointment", payload, ct).ConfigureAwait(false);
        }

        // ── HTTP core ─────────────────────────────────────────────────────────

        /// <summary>Like SendAsync but accepts a pre-built JSON string instead of an object.</summary>
        private async Task<T> SendRawAsync<T>(HttpMethod method, string relativePath, string jsonBody, CancellationToken ct)
        {
            var token = _tokenStore.GetToken();
            if (string.IsNullOrWhiteSpace(token)) throw new ProofioUnauthorizedException();

            var requestUrl = NormalizeBaseUrl(_tokenStore.GetApiBaseUrl()) + NormalizeRelativePath(relativePath);
            Logger.Info("HTTP-Call (raw): " + method + " " + requestUrl);
            Logger.Info("Payload: " + jsonBody);
            WriteDebugPayload("last-case-payload.json", jsonBody);

            using (var request = new HttpRequestMessage(method, requestUrl))
            {
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                request.Headers.Accept.Clear();
                request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                request.Content = new StringContent(jsonBody, Encoding.UTF8, "application/json");

                try
                {
                    using (var response = await SharedClient.Value.SendAsync(request, ct).ConfigureAwait(false))
                    {
                        var text = response.Content == null ? string.Empty
                            : await response.Content.ReadAsStringAsync().ConfigureAwait(false);

                        Logger.Info("HTTP-Antwort: " + (int)response.StatusCode + " " + response.ReasonPhrase);

                        if (response.StatusCode == HttpStatusCode.Unauthorized) throw new ProofioUnauthorizedException();
                        if (!response.IsSuccessStatusCode)
                            throw new HttpRequestException("Proofio API: " + (int)response.StatusCode + " " + response.ReasonPhrase + " – " + text);

                        if (typeof(T) == typeof(object) || string.IsNullOrWhiteSpace(text)) return default(T);
                        return FromJson<T>(text);
                    }
                }
                catch (ProofioUnauthorizedException) { throw; }
                catch (Exception ex) { Logger.Error("HTTP-Call fehlgeschlagen: " + requestUrl, ex); throw; }
            }
        }

        private async Task<T> SendAsync<T>(HttpMethod method, string relativePath, object payload, CancellationToken ct)
        {
            var token = _tokenStore.GetToken();
            if (string.IsNullOrWhiteSpace(token))
                throw new ProofioUnauthorizedException();

            var baseUrl    = NormalizeBaseUrl(_tokenStore.GetApiBaseUrl());
            var requestUrl = baseUrl + NormalizeRelativePath(relativePath);

            using (var request = new HttpRequestMessage(method, requestUrl))
            {
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                request.Headers.Accept.Clear();
                request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

                if (payload != null)
                {
                    var json = ToJson(payload);
                    Logger.Info("Proofio Request: " + method + " " + relativePath + " – " + json);

                    if (relativePath.IndexOf("/outlook/email",       StringComparison.OrdinalIgnoreCase) >= 0) WriteDebugPayload("last-email-payload.json",       json);
                    if (relativePath.IndexOf("/outlook/appointment", StringComparison.OrdinalIgnoreCase) >= 0) WriteDebugPayload("last-appointment-payload.json", json);
                    if (relativePath.IndexOf("/outlook/case",        StringComparison.OrdinalIgnoreCase) >= 0) WriteDebugPayload("last-case-payload.json",        json);

                    request.Content = new StringContent(json, Encoding.UTF8, "application/json");
                }

                try
                {
                    Logger.Info("HTTP-Call: " + method + " " + requestUrl);

                    using (var response = await SharedClient.Value.SendAsync(request, ct).ConfigureAwait(false))
                    {
                        var text = response.Content == null
                            ? string.Empty
                            : await response.Content.ReadAsStringAsync().ConfigureAwait(false);

                        Logger.Info("HTTP-Antwort: " + (int)response.StatusCode + " " + response.ReasonPhrase + " für " + requestUrl);

                        if (response.StatusCode == HttpStatusCode.Unauthorized)
                            throw new ProofioUnauthorizedException();

                        if (!response.IsSuccessStatusCode)
                            throw new HttpRequestException(
                                "Proofio API: " + (int)response.StatusCode + " " + response.ReasonPhrase + " – " + text);

                        if (typeof(T) == typeof(object) || string.IsNullOrWhiteSpace(text))
                            return default(T);

                        return FromJson<T>(text);
                    }
                }
                catch (ProofioUnauthorizedException) { throw; }
                catch (Exception ex)
                {
                    Logger.Error("HTTP-Call fehlgeschlagen: " + method + " " + requestUrl, ex);
                    throw;
                }
            }
        }

        // ── Helpers ───────────────────────────────────────────────────────────

        private static string NormalizeBaseUrl(string baseUrl)
        {
            if (string.IsNullOrWhiteSpace(baseUrl)) return TokenStore.DefaultApiBaseUrl;
            return baseUrl.Trim().TrimEnd('/');
        }

        private static string NormalizeRelativePath(string relativePath)
        {
            if (string.IsNullOrWhiteSpace(relativePath)) return string.Empty;
            relativePath = relativePath.Trim();
            return relativePath.StartsWith("/") ? relativePath : "/" + relativePath;
        }

        private static string ToJson(object value)
        {
            var serializer = new DataContractJsonSerializer(value.GetType());
            using (var ms = new MemoryStream())
            {
                serializer.WriteObject(ms, value);
                return Encoding.UTF8.GetString(ms.ToArray());
            }
        }

        private static T FromJson<T>(string json)
        {
            var serializer = new DataContractJsonSerializer(typeof(T));
            using (var ms = new MemoryStream(Encoding.UTF8.GetBytes(json)))
            {
                return (T)serializer.ReadObject(ms);
            }
        }

        private static void WriteDebugPayload(string fileName, string json)
        {
            try
            {
                var dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Proofio");
                Directory.CreateDirectory(dir);
                File.WriteAllText(Path.Combine(dir, fileName), json, Encoding.UTF8);
            }
            catch (Exception ex) { Logger.Error("Debug-Payload konnte nicht geschrieben werden.", ex); }
        }
    }
}
