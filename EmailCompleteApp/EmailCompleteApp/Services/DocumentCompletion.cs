using EmailCompleteApp.Models;
using NPOI.XWPF.UserModel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Threading.Tasks;
using System.Windows; 

namespace EmailCompleteApp.Services
{
    public class DocumentCompletion
    {
        private static DocumentCompletion? _instance;
        private static readonly object _lock = new object();

        private const int MaxParentDirectoryLevels = 6;
        private const string DocumentFolderName = "doc";
        private const string TemplateFileName = "comanda.docx";
        private const string GeneratedFolderName = "Generated";

        public static DocumentCompletion Instance
        {
            get
            {
                if (_instance == null)
                {
                    lock (_lock)
                    {
                        _instance ??= new DocumentCompletion();
                    }
                }
                return _instance;
            }
        }

        private DocumentCompletion() { }

        /// <summary>
        /// Public entry point to generate the document. ALL required data is passed explicitly via parameters.
        /// </summary>
        public async Task GenerateAndSendDocumentAsync(
            string numarComanda,
            string numarClient,
            string client,
            string tarif,
            int monedaIndex,
            int tipIndex,
            string transportator,
            string transportatorTarif,
            int transportatorMonedaIndex,
            int transportatorTipIndex,
            DateTime? dataIncarcare,
            DateTime? dataDescarcare,
            string produs,
            string cantitate,
            int tipAdrIndex,
            string clasa,
            string un,
            string numarInmatriculare,
            // Location (pickup) components
            string locatieIncarcareAddress,
            string locatieIncarcareName,
            string locatieIncarcareCity,
            string locatieIncarcareCode,
            // Location (delivery) components
            string locatieDescarcareAddress,
            string locatieDescarcareName,
            string locatieDescarcareCity,
            string locatieDescarcareCode,
            string termenPlata,
            string commentUser,
            // Option arrays (must be supplied by caller UI/ViewModel)
            string[] monedaOptions,
            string[] tipOptions,
            string[] tipAdrOptions
        )
        {
            try
            {
                string projectRoot = AppDomain.CurrentDomain.BaseDirectory;
                string projectDir = FindProjectDirectory(projectRoot);
                string docDir = Path.Combine(projectDir, DocumentFolderName);
                string templatePath = Path.Combine(docDir, TemplateFileName);

                if (!File.Exists(templatePath))
                {
                    ShowError($"No template found. Add '{TemplateFileName}' under: {docDir}", "Template Missing");
                    return;
                }

                string generatedDir = Path.Combine(docDir, GeneratedFolderName);
                Directory.CreateDirectory(generatedDir);

                string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH.mm.ss");
                string outputPath = Path.Combine(generatedDir, $"CAPAC+Comanda transport - {numarComanda}.docx");

                var replacements = await BuildReplacementDictionary(
                    numarComanda,
                    numarClient,
                    client,
                    tarif,
                    monedaIndex,
                    tipIndex,
                    transportator,
                    transportatorTarif,
                    transportatorMonedaIndex,
                    transportatorTipIndex,
                    dataIncarcare,
                    dataDescarcare,
                    produs,
                    cantitate,
                    tipAdrIndex,
                    clasa,
                    un,
                    numarInmatriculare,
                    locatieIncarcareAddress,
                    locatieIncarcareName,
                    locatieIncarcareCity,
                    locatieIncarcareCode,
                    locatieDescarcareAddress,
                    locatieDescarcareName,
                    locatieDescarcareCity,
                    locatieDescarcareCode,
                    termenPlata,
                    commentUser,
                    monedaOptions,
                    tipOptions,
                    tipAdrOptions
                );

                bool success = GenerateWordDocument(templatePath, outputPath, replacements);

                if (success)
                {
                    ShowSuccess($"DOCX generated.\n\nDOCX: {outputPath}", "Ready to Send");
                }
                else
                {
                    await TryOpenDocumentAsync(outputPath);
                }
            }
            catch (Exception ex)
            {
                ShowError($"Failed to generate document.\n\nError: {ex.Message}", "Error");
            }
        }

        /// <summary>
        /// Build dictionary of placeholder replacements for document generation
        /// </summary>
        private Task<Dictionary<string, string>> BuildReplacementDictionary(
            string numarComanda,
            string numarClient,
            string client,
            string tarif,
            int monedaIndex,
            int tipIndex,
            string transportator,
            string transportatorTarif,
            int transportatorMonedaIndex,
            int transportatorTipIndex,
            DateTime? dataIncarcare,
            DateTime? dataDescarcare,
            string produs,
            string cantitate,
            int tipAdrIndex,
            string clasa,
            string un,
            string numarInmatriculare,
            //incarcare
            string locatieIncarcareAddress,
            string locatieIncarcareName,
            string locatieIncarcareCity,
            string locatieIncarcareCode,
            //descarcare
            string locatieDescarcareAddress,
            string locatieDescarcareName,
            string locatieDescarcareCity,
            string locatieDescarcareCode,
            string termenPlata,
            string commentUser,
            string[] monedaOptions,
            string[] tipOptions,
            string[] tipAdrOptions
        )
        {
            DateTime today = DateTime.Today;

            return Task.FromResult(new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "DataAzi", today.ToString("dd.MM.yyyy") },
                { "NumarComanda", numarComanda?.Trim() ?? string.Empty },
                { "NumarClient", numarClient?.Trim() ?? string.Empty },
                { "ClientNume", client?.Trim() ?? string.Empty },
                { "ClientTarif", tarif?.Trim() ?? string.Empty },
                { "ClientMoneda", GetOptionValue(monedaIndex, monedaOptions) },
                { "ClientTip", GetOptionValue(tipIndex, tipOptions) },
                { "TransportatorNume", transportator?.Trim() ?? string.Empty },
                { "TransportatorTarif", transportatorTarif?.Trim() ?? string.Empty },
                { "TransportatorMoneda", GetOptionValue(transportatorMonedaIndex, monedaOptions) },
                { "TransportatorTip", GetOptionValue(transportatorTipIndex, tipOptions) },
                { "DataIncarcare", dataIncarcare?.ToString("dd/MM/yyyy") ?? string.Empty },
                { "DataDescarcare", dataDescarcare?.ToString("dd/MM/yyyy") ?? string.Empty },
                { "Produs", produs?.Trim() ?? string.Empty },
                { "CantitateComanda", cantitate?.Trim() ?? string.Empty },
                { "TipADR", GetOptionValue(tipAdrIndex, tipAdrOptions) },
                { "Clasa", clasa?.Trim() ?? string.Empty },
                { "UnInput", un?.Trim() ?? string.Empty },
                { "NumarInmatriculare", numarInmatriculare?.Trim().ToUpper() ?? string.Empty },
                { "AdresaIncarcare", locatieIncarcareAddress?.Trim() ?? string.Empty },
                { "AddresaIncarcareName", locatieIncarcareName?.Trim() ?? string.Empty },
                { "AddresaIncarcareCity", locatieIncarcareCity?.Trim() ?? string.Empty },
                { "AddresaIncarcareCityCode", locatieIncarcareCode?.Trim() ?? string.Empty },
                { "AdresaDescarcare", locatieDescarcareAddress?.Trim() ?? string.Empty },
                { "AddresaDescarcareName", locatieDescarcareName?.Trim() ?? string.Empty },
                { "AddresaDescarcareCity", locatieDescarcareCity?.Trim() ?? string.Empty },
                { "AddresaDescarcareCityCode", locatieDescarcareCode?.Trim() ?? string.Empty },
                { "TermenPlata", termenPlata?.Trim() ?? string.Empty },
                { "Comments", commentUser?.Trim() ?? string.Empty }
            });
        }

        private static string GetOptionValue(int index, string[] options)
        {
            return options != null && index >= 0 && index < options.Length ? options[index] : string.Empty;
        }

        #region Word Document Processing

        private static bool GenerateWordDocument(string templatePath, string outputPath, Dictionary<string, string> replacements)
        {
            try
            {
                using (var fs = new FileStream(templatePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var document = new XWPFDocument(fs))
                {
                    ReplaceInDocument(document, replacements);

                    using (var outFs = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
                    {
                        document.Write(outFs);
                    }
                }
                return true;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error processing Word document: {ex.Message}");
                return false;
            }
        }

        private static void ReplaceInDocument(XWPFDocument document, Dictionary<string, string> replacements)
        {
            foreach (var paragraph in document.Paragraphs)
            {
                ReplaceInParagraph(paragraph, replacements);
            }
            foreach (var table in document.Tables)
            {
                ReplaceInTable(table, replacements);
            }
            foreach (var header in document.HeaderList)
            {
                foreach (var paragraph in header.Paragraphs)
                {
                    ReplaceInParagraph(paragraph, replacements);
                }
                foreach (var table in header.Tables)
                {
                    ReplaceInTable(table, replacements);
                }
            }
            foreach (var footer in document.FooterList)
            {
                foreach (var paragraph in footer.Paragraphs)
                {
                    ReplaceInParagraph(paragraph, replacements);
                }
                foreach (var table in footer.Tables)
                {
                    ReplaceInTable(table, replacements);
                }
            }
        }

        private static void ReplaceInTable(NPOI.XWPF.UserModel.XWPFTable table, Dictionary<string, string> replacements)
        {
            foreach (var row in table.Rows)
            {
                foreach (var cell in row.GetTableCells())
                {
                    foreach (var paragraph in cell.Paragraphs)
                    {
                        ReplaceInParagraph(paragraph, replacements);
                    }
                    foreach (var innerTable in cell.Tables)
                    {
                        ReplaceInTable(innerTable, replacements);
                    }
                }
            }
        }

        private static void ReplaceInParagraph(XWPFParagraph paragraph, Dictionary<string, string> replacements)
        {
            for (int i = 0; i < paragraph.Runs.Count; i++)
            {
                var run = paragraph.Runs[i];
                string? text = run.Text;
                if (string.IsNullOrEmpty(text)) continue;

                string modifiedText = text;
                foreach (var replacement in replacements)
                {
                    if (string.IsNullOrEmpty(replacement.Key)) continue;
                    modifiedText = modifiedText.Replace(replacement.Key, replacement.Value ?? string.Empty, StringComparison.OrdinalIgnoreCase);
                }
                if (!string.Equals(text, modifiedText, StringComparison.Ordinal))
                {
                    run.SetText(modifiedText, 0);
                }
            }
        }

        #endregion

        #region File System Helpers
        private static string FindProjectDirectory(string startPath)
        {
            string? current = startPath;
            for (int i = 0; i < MaxParentDirectoryLevels && current != null; i++)
            {
                string candidate = Path.Combine(current, DocumentFolderName);
                if (Directory.Exists(candidate))
                {
                    return current;
                }
                current = Directory.GetParent(current)?.FullName;
            }
            return startPath;
        }
        #endregion

        private async Task TryOpenDocumentAsync(string documentPath)
        {
            try
            {
                await Task.Run(() => Process.Start(new ProcessStartInfo(documentPath) { UseShellExecute = true }));
                ShowSuccess($"DOCX generated.\n\nDOCX: {documentPath}\n\nOpened directly.", "Success");
            }
            catch (Exception openEx)
            {
                ShowWarning($"DOCX generated but could not be opened.\n\nDOCX: {documentPath}\nError: {openEx.Message}", "Open Failed");
            }
        }

        private static void ShowError(string message, string title)
        {
            Application.Current.Dispatcher.Invoke(() => MessageBox.Show(message, title, MessageBoxButton.OK, MessageBoxImage.Error));
        }

        private static void ShowSuccess(string message, string title)
        {
            Application.Current.Dispatcher.Invoke(() => MessageBox.Show(message, title, MessageBoxButton.OK, MessageBoxImage.Information));
        }

        private static void ShowWarning(string message, string title)
        {
            Application.Current.Dispatcher.Invoke(() => MessageBox.Show(message, title, MessageBoxButton.OK, MessageBoxImage.Warning));
        }
    }
}
