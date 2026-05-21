using System;
using System.Runtime.InteropServices;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace ProofioAddIn.Services
{
    /// <summary>
    /// Stamps a green Outlook category on a MailItem after it has been
    /// successfully filed in Proofio. The category appears as a coloured
    /// badge directly in the inbox / folder view.
    /// </summary>
    public static class MailCategoryService
    {
        // ── Constants ────────────────────────────────────────────────────────

        /// <summary>Display name of the category shown in Outlook.</summary>
        public const string CategoryName = "In Proofio abgelegt";

        /// <summary>
        /// Outlook colour constant for green.
        /// olCategoryColorDarkGreen = 6 (dark green, clearly visible)
        /// olCategoryColorGreen     = 5 (lighter green)
        /// </summary>
        private const Outlook.OlCategoryColor CategoryColor =
            Outlook.OlCategoryColor.olCategoryColorDarkGreen;

        // ── Public API ───────────────────────────────────────────────────────

        /// <summary>
        /// Ensures the Proofio category exists in the user's master category
        /// list, then applies it to <paramref name="mail"/>.
        /// Safe to call from any thread context; all COM access stays on the
        /// calling thread (must be the UI / COM thread).
        /// </summary>
        public static void Apply(Outlook.MailItem mail)
        {
            if (mail == null) return;

            try
            {
                EnsureCategoryExists(mail);
                AppendCategory(mail);
            }
            catch (COMException ex)
            {
                Logger.Error("MailCategoryService: Kategorie setzen fehlgeschlagen.", ex);
            }
            catch (Exception ex)
            {
                Logger.Error("MailCategoryService: Unerwarteter Fehler.", ex);
            }
        }

        // ── Internals ────────────────────────────────────────────────────────

        /// <summary>
        /// Registers the category in Outlook's master list if it is not
        /// already present, using the chosen colour.
        /// </summary>
        private static void EnsureCategoryExists(Outlook.MailItem mail)
        {
            Outlook.Categories categories = null;
            try
            {
                // Categories live on the Store (mailbox) the item belongs to.
                var store  = mail.Parent as Outlook.Folder;
                categories = store?.Store?.Categories
                          ?? Globals.ThisAddIn.Application.Session.Categories;

                foreach (Outlook.Category cat in categories)
                {
                    if (string.Equals(cat.Name, CategoryName,
                            StringComparison.OrdinalIgnoreCase))
                        return; // already exists
                }

                // Create it once
                categories.Add(CategoryName, CategoryColor);
                Logger.Info("MailCategoryService: Kategorie '" + CategoryName + "' angelegt.");
            }
            finally
            {
                if (categories != null) Marshal.ReleaseComObject(categories);
            }
        }

        /// <summary>
        /// Appends <see cref="CategoryName"/> to the mail's categories string
        /// without removing any existing categories, then saves the change.
        /// </summary>
        private static void AppendCategory(Outlook.MailItem mail)
        {
            var existing = mail.Categories ?? string.Empty;

            // Avoid duplicates
            if (existing.IndexOf(CategoryName, StringComparison.OrdinalIgnoreCase) >= 0)
                return;

            mail.Categories = string.IsNullOrWhiteSpace(existing)
                ? CategoryName
                : existing + "; " + CategoryName;

            mail.Save();

            Logger.Info("MailCategoryService: Kategorie auf Mail gesetzt – " + mail.Subject);
        }
    }
}
