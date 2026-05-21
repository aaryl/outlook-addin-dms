using System;
using System.Configuration;
using Microsoft.Win32;

namespace ProofioAddIn.Services
{
    public sealed class TokenStore
    {
        private const string RegistryPath = @"Software\Proofio";
        private const string TokenName = "Token";
        private const string ApiBaseUrlName = "ApiBaseUrl";
        private const string SendFilingModeName = "SendFilingMode";

        public const string DefaultApiBaseUrl = "https://gutachtenpilot.lovable.app/api/public/v1";

        public const string SendFilingModeNever = "Never";
        public const string SendFilingModeAlways = "Always";
        public const string SendFilingModeAsk = "Ask";

        public string GetToken()
        {
            return ReadString(TokenName);
        }

        public string GetApiBaseUrl()
        {
            var stored = ReadString(ApiBaseUrlName);

            if (!string.IsNullOrWhiteSpace(stored))
            {
                return NormalizeBaseUrl(stored);
            }

            var configured = ConfigurationManager.AppSettings["Proofio.DefaultApiBaseUrl"];
            return NormalizeBaseUrl(string.IsNullOrWhiteSpace(configured) ? DefaultApiBaseUrl : configured);
        }

        public string GetSendFilingMode()
        {
            var value = ReadString(SendFilingModeName);

            if (string.Equals(value, SendFilingModeAlways, StringComparison.OrdinalIgnoreCase))
            {
                return SendFilingModeAlways;
            }

            if (string.Equals(value, SendFilingModeAsk, StringComparison.OrdinalIgnoreCase))
            {
                return SendFilingModeAsk;
            }

            return SendFilingModeNever;
        }

        public void Save(string token, string apiBaseUrl)
        {
            Save(token, apiBaseUrl, GetSendFilingMode());
        }

        public void Save(string token, string apiBaseUrl, string sendFilingMode)
        {
            using (var root = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, RegistryView.Registry64))
            using (var key = root.CreateSubKey(RegistryPath))
            {
                if (key == null)
                {
                    throw new InvalidOperationException("Registry-Schlüssel konnte nicht erstellt werden.");
                }

                key.SetValue(TokenName, token ?? string.Empty, RegistryValueKind.String);
                key.SetValue(ApiBaseUrlName, NormalizeBaseUrl(apiBaseUrl), RegistryValueKind.String);
                key.SetValue(SendFilingModeName, NormalizeSendFilingMode(sendFilingMode), RegistryValueKind.String);
            }
        }

        private static string ReadString(string name)
        {
            using (var root = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, RegistryView.Registry64))
            using (var key = root.OpenSubKey(RegistryPath, false))
            {
                return key == null ? null : key.GetValue(name) as string;
            }
        }

        private static string NormalizeBaseUrl(string value)
        {
            value = string.IsNullOrWhiteSpace(value) ? DefaultApiBaseUrl : value.Trim();
            return value.TrimEnd('/');
        }

        private static string NormalizeSendFilingMode(string value)
        {
            if (string.Equals(value, SendFilingModeAlways, StringComparison.OrdinalIgnoreCase))
            {
                return SendFilingModeAlways;
            }

            if (string.Equals(value, SendFilingModeAsk, StringComparison.OrdinalIgnoreCase))
            {
                return SendFilingModeAsk;
            }

            return SendFilingModeNever;
        }
    }
}
