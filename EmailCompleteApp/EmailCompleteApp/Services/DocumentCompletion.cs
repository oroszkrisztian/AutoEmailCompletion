using EmailCompleteApp.Models;
using NPOI.XWPF.UserModel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
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
        private const string EmailFolderName = "Email"; 

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
        /// Returns the output path of the generated document.
        /// </summary>
        public async Task<string?> GenerateAndSendDocumentAsync(
            string numarComanda,
            string numarClient,
            string client,
            string contact,
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
            string locatieIncarcareCountryCode,
            string locatieIncarcarePostalCode,
            string locatieIncarcareCounty,
            // Location (delivery) components
            string locatieDescarcareAddress,
            string locatieDescarcareName,
            string locatieDescarcareCity,
            string locatieDescarcareCountryCode,
            string locatieDescarcarePostalCode,
            string locatieDescarcareCounty,
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
                    return null;
                }

                string generatedDir = Path.Combine(docDir, GeneratedFolderName);
                Directory.CreateDirectory(generatedDir);

                string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH.mm.ss");
                string outputPath = Path.Combine(generatedDir, $"Comanda {numarComanda} {locatieIncarcareCity} - {locatieDescarcareCity}.docx");

                var replacements = await BuildReplacementDictionary(
                    numarComanda,
                    numarClient,
                    client,
                    contact,
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
                    //incarcare
                    locatieIncarcareAddress,
                    locatieIncarcareName,
                    locatieIncarcareCity,
                    locatieIncarcareCountryCode,
                    locatieIncarcarePostalCode,
                    locatieIncarcareCounty,
                    //descarcare
                    locatieDescarcareAddress,
                    locatieDescarcareName,
                    locatieDescarcareCity,
                    locatieDescarcareCountryCode,
                    locatieDescarcarePostalCode,
                    locatieDescarcareCounty,
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
                    return outputPath;
                }
                
                return null;
            }
            catch (Exception ex)
            {
                ShowError($"Failed to generate document.\n\nError: {ex.Message}", "Error");
                return null;
            }
        }

        /// <summary>
        /// Build dictionary of placeholder replacements for document generation
        /// </summary>
        private Task<Dictionary<string, string>> BuildReplacementDictionary(
            string numarComanda,
            string numarClient,
            string client,
            string contact,
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
            string locatieIncarcareCountryCode,
            string locatieIncarcarePostalCode,
            string locatieIncarcareCounty,
            //descarcare
            string locatieDescarcareAddress,
            string locatieDescarcareName,
            string locatieDescarcareCity,
            string locatieDescarcareCountryCode,
            string locatieDescarcarePostalCode,
            string locatieDescarcareCounty,
            string termenPlata,
            string commentUser,
            string[] monedaOptions,
            string[] tipOptions,
            string[] tipAdrOptions
        )
        {
            DateTime today = DateTime.Today;

            int tipCLientIndex = tipIndex;
            int tipTransportatorIndex = transportatorTipIndex;
            string tvaClient = tipCLientIndex == 0 ? "+ " + GetOptionValue(tipCLientIndex, tipOptions) : GetOptionValue(tipCLientIndex, tipOptions);
            string tvaTransportator = tipTransportatorIndex == 0 ? "+ " + GetOptionValue(tipTransportatorIndex, tipOptions) : GetOptionValue(tipTransportatorIndex, tipOptions);



            return Task.FromResult(new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "DataAzi", today.ToString("dd.MM.yyyy") },
                { "NumarComanda", numarComanda?.Trim() ?? string.Empty },
                { "NumarClient", numarClient?.Trim() ?? string.Empty },
                { "ClientNume", client?.Trim() ?? string.Empty },
                { "ContactPers", contact?.Trim() ?? string.Empty },
                { "ClientTarif", tarif?.Trim() ?? string.Empty },
                { "ClientMoneda", GetOptionValue(monedaIndex, monedaOptions) },
                { "ClientTip", tvaClient },
                { "TransportatorNume", transportator?.Trim() ?? string.Empty },
                { "TransportatorTarif", transportatorTarif?.Trim() ?? string.Empty },
                { "TransportatorMoneda", GetOptionValue(transportatorMonedaIndex, monedaOptions) },
                { "TransportatorTip", tvaTransportator },
                { "DataIncarcare", dataIncarcare?.ToString("dd/MM/yyyy") ?? string.Empty },
                { "DataDescarcare", dataDescarcare?.ToString("dd/MM/yyyy") ?? string.Empty },
                { "Produs", produs?.Trim() ?? string.Empty },
                { "CantitateComanda", cantitate?.Trim() ?? string.Empty },
                { "TipADR", GetOptionValue(tipAdrIndex, tipAdrOptions) },
                { "Clasa", clasa?.Trim() ?? string.Empty },
                { "UserUnInput", un?.Trim() ?? string.Empty },
                { "NumarInmatriculare", numarInmatriculare?.Trim().ToUpper() ?? string.Empty },
                { "LocatieIncarcareAddress", locatieIncarcareAddress?.Trim() ?? string.Empty },
                { "LocatieIncarcareName", locatieIncarcareName?.Trim() ?? string.Empty },
                { "LocatieIncarcareCity", locatieIncarcareCity?.Trim() ?? string.Empty },
                { "LocatieIncarcareCountryCode", locatieIncarcareCountryCode?.Trim() ?? string.Empty },
                { "LocatieIncarcarePostalCode", locatieIncarcarePostalCode?.Trim() ?? string.Empty },
                { "LocatieIncarcareCounty", locatieIncarcareCounty?.Trim() ?? string.Empty },
                { "LocatieDescarcareAddress", locatieDescarcareAddress?.Trim() ?? string.Empty },
                { "LocatieDescarcareName", locatieDescarcareName?.Trim() ?? string.Empty },
                { "LocatieDescarcareCity", locatieDescarcareCity?.Trim() ?? string.Empty },
                { "LocatieDescarcareCountryCode", locatieDescarcareCountryCode?.Trim() ?? string.Empty },
                { "LocatieDescarcarePostalCode", locatieDescarcarePostalCode?.Trim() ?? string.Empty },
                { "LocatieDescarcareCounty", locatieDescarcareCounty?.Trim() ?? string.Empty },
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
                bool termenPlataReplaced = false;
                foreach (var replacement in replacements)
                {
                    if (string.IsNullOrEmpty(replacement.Key)) continue;
                    // Check if termenPlata is being replaced
                    if (replacement.Key.Equals("termenPlata", StringComparison.OrdinalIgnoreCase) && text.Contains(replacement.Key, StringComparison.OrdinalIgnoreCase))
                    {
                        termenPlataReplaced = true;
                    }
                    modifiedText = modifiedText.Replace(replacement.Key, replacement.Value ?? string.Empty, StringComparison.OrdinalIgnoreCase);
                }
                if (!string.Equals(text, modifiedText, StringComparison.Ordinal))
                {
                    run.SetText(modifiedText, 0);
                    if (termenPlataReplaced)
                    {
                        run.SetColor("FF0000"); 
                    }
                }
            }
        }

        #endregion

        #region Thunderbird Integration

        /// <summary>
        /// Opens Thunderbird email client with multiple documents attached to a new email
        /// </summary>
        private static void OpenThunderbirdWithAttachments(params string[] attachmentPaths)
        {
            try
            {
                // Validate all files exist
                foreach (var path in attachmentPaths)
                {
                    if (!File.Exists(path))
                    {
                        ShowWarning($"Document not found: {path}", "File Not Found");
                        return;
                    }
                }

                string[] possibleThunderbirdPaths = new[]
                {
                    @"C:\Program Files\Mozilla Thunderbird\thunderbird.exe",
                    @"C:\Program Files (x86)\Mozilla Thunderbird\thunderbird.exe",
                    Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), @"Mozilla Thunderbird\thunderbird.exe"),
                    Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), @"Mozilla Thunderbird\thunderbird.exe")
                };

                string? thunderbirdPath = null;
                foreach (var path in possibleThunderbirdPaths)
                {
                    if (File.Exists(path))
                    {
                        thunderbirdPath = path;
                        break;
                    }
                }

                if (string.IsNullOrEmpty(thunderbirdPath))
                {
                    ShowWarning(
                        "Thunderbird not found in default installation locations.\n\n" +
                        "Please open Thunderbird manually and attach the documents.",
                        "Thunderbird Not Found");
                    return;
                }

                // Build comma-separated list of file URIs for multiple attachments
                var fileUris = attachmentPaths.Select(path => new Uri(path).AbsoluteUri);
                string attachmentList = string.Join(",", fileUris);

                var startInfo = new ProcessStartInfo
                {
                    FileName = thunderbirdPath,
                    Arguments = $"-compose \"attachment='{attachmentList}'\"",
                    UseShellExecute = false
                };

                Process.Start(startInfo);
                Debug.WriteLine($"✅ Opened Thunderbird with {attachmentPaths.Length} attachment(s)");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"❌ Error opening Thunderbird: {ex.Message}");
                ShowWarning(
                    $"Could not open Thunderbird automatically.\n\n" +
                    $"Error: {ex.Message}\n\n" +
                    $"Please open Thunderbird manually.",
                    "Thunderbird Error");
            }
        }

        /// <summary>
        /// Opens Thunderbird email client with the specified document attached to a new email
        /// </summary>
        private static void OpenThunderbirdWithAttachment(string attachmentPath)
        {
            OpenThunderbirdWithAttachments(attachmentPath);
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

        /// <summary>
        /// Generates page2.doc in the Email folder, returns the path.
        /// Files will be cleaned up when user starts a new order.
        /// </summary>
        public async Task<string?> GenerateAndSendPage2DocAsync(
            string templatePath,
            Dictionary<string, string> replacements,
            string numarComanda
        )
        {
            try
            {
                // Check if template exists
                if (!File.Exists(templatePath))
                {
                    Debug.WriteLine($"❌ page2.docx template not found at: {templatePath}");
                    return null;
                }

                // Find project directory and Email folder
                string projectRoot = AppDomain.CurrentDomain.BaseDirectory;
                string projectDir = FindProjectDirectory(projectRoot);
                string docDir = Path.Combine(projectDir, DocumentFolderName);
                string emailDir = Path.Combine(docDir, EmailFolderName);
                
                // Create Email folder if it doesn't exist
                Directory.CreateDirectory(emailDir);

                // Create file with proper naming: Comanda {number}.doc in Email folder
                string outputFileName = $"Comanda {numarComanda}.doc";
                string outputPath = Path.Combine(emailDir, outputFileName);

                Debug.WriteLine($"📝 Generating page2.doc at: {outputPath}");

                // Generate the DOC file (synchronous to ensure it's written before returning)
                bool success = GenerateWordDocument(templatePath, outputPath, replacements);

                if (!success)
                {
                    Debug.WriteLine($"❌ Failed to generate Word document");
                    return null;
                }

                // Verify file exists and has content
                if (!File.Exists(outputPath))
                {
                    Debug.WriteLine($"❌ File was not created: {outputPath}");
                    return null;
                }

                var fileInfo = new FileInfo(outputPath);
                if (fileInfo.Length == 0)
                {
                    Debug.WriteLine($"❌ File is empty: {outputPath}");
                    return null;
                }

                Debug.WriteLine($"✅ Generated page2.doc at: {outputPath} (Size: {fileInfo.Length} bytes)");
                
                // Small delay to ensure Windows has fully released file handles
                await Task.Delay(200);

                return outputPath;
            }
            catch (Exception ex)
            {
                ShowError($"Failed to generate page2.doc.\n\nError: {ex.Message}", "Error");
                Debug.WriteLine($"❌ Exception in GenerateAndSendPage2DocAsync: {ex}");
                return null;
            }
        }

        /// <summary>
        /// Opens Thunderbird with the specified attachments.
        /// Files in Email folder will be cleaned up when the next order is created.
        /// </summary>
        public void OpenThunderbirdAndCleanup(string[] attachments, string[] tempFilesToDelete = null)
        {
            OpenThunderbirdWithAttachments(attachments);
            
            // Note: tempFilesToDelete parameter is kept for backward compatibility but no longer used.
            // Files in Email folder will be cleaned up automatically when user starts a new order.
            Debug.WriteLine($"📧 Thunderbird opened with {attachments.Length} attachment(s)");
            Debug.WriteLine($"ℹ️ Files will be automatically cleaned up when next order is created");
        }
    }
}
