using Microsoft.Win32;
using NPOI.XWPF.UserModel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Threading;

namespace EmailCompleteApp.Pages
{
    public partial class ComandaTransport : UserControl
    {
        public ComandaTransport()
        {
            InitializeComponent();

            
            MonedaComboBox.SelectedIndex = 0;
            TipComboBox.SelectedIndex = 0;
            if (TransportatorMonedaComboBox != null) TransportatorMonedaComboBox.SelectedIndex = 0;
            if (TransportatorTipComboBox != null) TransportatorTipComboBox.SelectedIndex = 0;
            if (TipAdrComboBox != null) TipAdrComboBox.SelectedIndex = 0;

            // Set default dates (if present)
            if (DataIncarcareDatePicker != null) DataIncarcareDatePicker.SelectedDate = DateTime.Today;
            if (DataDescarcareDatePicker != null) DataDescarcareDatePicker.SelectedDate = DateTime.Today.AddDays(1);

            // Handle text box validation
            var textBoxes = new[] {
                NumarComandaTextBox, ClientTextBox, TarifTextBox, PrimitTextBox,
                TransportatorTextBox, TransportatorTarifTextBox, OferitTextBox,
                ProdusTextBox, CantitateTextBox, ClasaTextBox, UMTextBox,
                MaxDaysTextBox, NumarInmatriculareTextBox, LocatieIncarcareTextBox, LocatieDescarcareTextBox
            };
            
        }

       

        private async void OnSendClick(object sender, System.Windows.RoutedEventArgs e)
        {
            try
            {
                string projectRoot = AppDomain.CurrentDomain.BaseDirectory;

                // Navigate up from bin/Debug/... to project folder if running from build output
                string FindProjectDirWithDoc(string start)
                {
                    string? current = start;
                    for (int i = 0; i < 6 && current != null; i++)
                    {
                        string candidate = Path.Combine(current, "doc");
                        if (Directory.Exists(candidate))
                        {
                            return current;
                        }
                        current = Directory.GetParent(current)?.FullName;
                    }
                    return start;
                }

                string projectDir = FindProjectDirWithDoc(projectRoot);
                string docDir = Path.Combine(projectDir, "doc");

                // Use the pre-merged template
                string mergedTemplatePath = Path.Combine(docDir, "comanda.docx");

                string generatedDir = Path.Combine(docDir, "Generated");
                Directory.CreateDirectory(generatedDir);

                // Human-readable, Windows-safe timestamp (no colons)
                string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH.mm.ss");
                var replacements = BuildCombinedReplacements();

                string mergedOutputPath = Path.Combine(generatedDir, $"CAPAC+Comanda transport - {timestamp}.docx");

                // Check template exists
                if (!File.Exists(mergedTemplatePath))
                {
                    MessageBox.Show($"No template found. Add 'comanda.docx' under: {docDir}",
                                  "Template Missing", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                // Generate the document from the pre-merged template
                GenerateWordDocumentFromTemplate(mergedTemplatePath, mergedOutputPath, replacements);

                // Show loading dialog while preparing email
                var ownerWindow = Window.GetWindow(this);
                var loading = new LoadingWindow();
                if (ownerWindow != null) loading.Owner = ownerWindow;
                loading.Show();
                await Task.Delay(50); // allow UI to render

                // Create Outlook email with DOCX attached (if Outlook available)
                bool emailCreated = false;
                try
                {
                    emailCreated = await Task.Run(() => CreateOutlookEmailWithAttachment(mergedOutputPath));
                }
                catch (Exception mailEx)
                {
                    Debug.WriteLine($"Email creation failed: {mailEx}");
                    emailCreated = false;
                }

                try { loading.Close(); } catch { }

                if (emailCreated)
                {
                    MessageBox.Show($"DOCX generated and email draft opened.\n\nDOCX: {mergedOutputPath}",
                                    "Ready to Send", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                else
                {
                    // Open the document directly if email creation failed
                    try
                    {
                        Process.Start(new ProcessStartInfo(mergedOutputPath) { UseShellExecute = true });
                        MessageBox.Show($"DOCX generated.\n\nDOCX: {mergedOutputPath}\n\nOutlook not found or could not open email. The document has been opened directly.",
                                        "Success (Manual Email)", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                    catch (Exception openEx)
                    {
                        MessageBox.Show($"DOCX generated but could not be opened.\n\nDOCX: {mergedOutputPath}\nError: {openEx.Message}",
                                        "Success (Manual Open Failed)", MessageBoxButton.OK, MessageBoxImage.Warning);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to generate document.\n\nError: {ex.Message}",
                              "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private Dictionary<string, string> BuildCombinedReplacements()
        {
            string Get(string? s) => s?.Trim() ?? string.Empty;

            var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "Numar Comanda", NumarComandaTextBox.Text?.Trim() ?? string.Empty },
                { "Client", ClientTextBox.Text?.Trim() ?? string.Empty },
                { "Tarif", TarifTextBox.Text?.Trim() ?? string.Empty },
                { "Primit", PrimitTextBox.Text?.Trim() ?? string.Empty },
                { "Moneda", (MonedaComboBox.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? string.Empty },
                { "Tip", (TipComboBox.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? string.Empty }
            };

            return map;
        }

        private string BuildMaxDays()
        {
            return string.Empty;
        }

        private static void GenerateWordDocumentFromTemplate(string templatePath, string outputPath, Dictionary<string, string> replacements)
        {
            try
            {
                using (var fs = new FileStream(templatePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var document = new XWPFDocument(fs))
                {
                    // Replace in body paragraphs
                    foreach (var paragraph in document.Paragraphs)
                    {
                        ReplaceInParagraph(paragraph, replacements);
                    }

                    // Replace in tables
                    foreach (var table in document.Tables)
                    {
                        ReplaceInTable(table, replacements);
                    }

                    // Replace in headers
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

                    // Replace in footers
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

                    // Force all text color to black
                    ForceDocumentTextColorBlack(document);

                    using (var outFs = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
                    {
                        document.Write(outFs);
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Error occurred while processing Word document: {ex.Message}", ex);
            }
        }

        private static void ReplaceInTable(XWPFTable table, Dictionary<string, string> replacements)
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
            string originalParagraphText = paragraph.Text;

            var runs = paragraph.Runs;
            bool anyRunChanged = false;
            if (runs != null)
            {
                for (int i = 0; i < runs.Count; i++)
                {
                    string? text = runs[i].ToString();
                    if (string.IsNullOrEmpty(text))
                        continue;

                    string replaced = ReplaceAll(text, replacements);
                    if (!string.Equals(text, replaced, StringComparison.Ordinal))
                    {
                        runs[i].SetText(replaced, 0);
                        anyRunChanged = true;
                    }
                }
            }

           
            if (!anyRunChanged)
            {
                string newParaText = ReplaceAll(originalParagraphText, replacements);
                if (!string.Equals(originalParagraphText, newParaText, StringComparison.Ordinal))
                {
                    for (int i = paragraph.Runs.Count - 1; i >= 0; i--)
                    {
                        paragraph.RemoveRun(i);
                    }
                    var run = paragraph.CreateRun();
                    run.SetText(newParaText);
                }
            }
        }

        private static string ReplaceAll(string input, Dictionary<string, string> replacements)
        {
            string output = input;
            foreach (var kvp in replacements)
            {
                if (string.IsNullOrEmpty(kvp.Key)) continue;
                output = output.Replace(kvp.Key, kvp.Value ?? string.Empty, StringComparison.OrdinalIgnoreCase);
            }
            return output;
        }

        private static void ForceDocumentTextColorBlack(XWPFDocument document)
        {
            var black = "000000";

            void SetRunsBlack(IEnumerable<XWPFRun> runs)
            {
                foreach (var run in runs)
                {
                    try
                    {
                        run.SetColor(black);
                    }
                    catch { }
                }
            }

            foreach (var paragraph in document.Paragraphs)
            {
                SetRunsBlack(paragraph.Runs);
            }

            foreach (var table in document.Tables)
            {
                foreach (var row in table.Rows)
                {
                    foreach (var cell in row.GetTableCells())
                    {
                        foreach (var paragraph in cell.Paragraphs)
                        {
                            SetRunsBlack(paragraph.Runs);
                        }
                        foreach (var innerTable in cell.Tables)
                        {
                            foreach (var innerRow in innerTable.Rows)
                            {
                                foreach (var innerCell in innerRow.GetTableCells())
                                {
                                    foreach (var innerPara in innerCell.Paragraphs)
                                    {
                                        SetRunsBlack(innerPara.Runs);
                                    }
                                }
                            }
                        }
                    }
                }
            }

            // Headers
            foreach (var header in document.HeaderList)
            {
                foreach (var paragraph in header.Paragraphs)
                {
                    SetRunsBlack(paragraph.Runs);
                }
                foreach (var table in header.Tables)
                {
                    foreach (var row in table.Rows)
                    {
                        foreach (var cell in row.GetTableCells())
                        {
                            foreach (var paragraph in cell.Paragraphs)
                            {
                                SetRunsBlack(paragraph.Runs);
                            }
                        }
                    }
                }
            }

            // Footers
            foreach (var footer in document.FooterList)
            {
                foreach (var paragraph in footer.Paragraphs)
                {
                    SetRunsBlack(paragraph.Runs);
                }
                foreach (var table in footer.Tables)
                {
                    foreach (var row in table.Rows)
                    {
                        foreach (var cell in row.GetTableCells())
                        {
                            foreach (var paragraph in cell.Paragraphs)
                            {
                                SetRunsBlack(paragraph.Runs);
                            }
                        }
                    }
                }
            }
        }

        private static bool CreateOutlookEmailWithAttachment(string attachmentPath)
        {
            if (!File.Exists(attachmentPath))
                throw new FileNotFoundException("Attachment not found", attachmentPath);

            Type? outlookType = Type.GetTypeFromProgID("Outlook.Application");
            if (outlookType == null)
            {
                return false; 
            }

            object? outlookApp = null;
            object? mailItem = null;
            try
            {
                outlookApp = Activator.CreateInstance(outlookType);
                if (outlookApp == null) return false;

                // 0 => olMailItem
                mailItem = outlookType
                    .GetMethod("CreateItem")?
                    .Invoke(outlookApp, new object[] { 0 });
                if (mailItem == null) return false;

                var mailType = mailItem.GetType();
                mailType.GetProperty("Subject")?.SetValue(mailItem, "Comanda transport");
                mailType.GetProperty("Body")?.SetValue(mailItem, "Va rugam gasiti atasat documentul in format DOCX.");

                var attachments = mailType.GetProperty("Attachments")?.GetValue(mailItem);
                var attachmentsType = attachments?.GetType();
                attachmentsType?.GetMethod("Add")?.Invoke(attachments, new object[] { attachmentPath });

                // Display the email for user to review/send
                mailType.GetMethod("Display", new[] { typeof(object) })?.Invoke(mailItem, new object?[] { false });
                return true;
            }
            finally
            {
                if (mailItem != null) Marshal.FinalReleaseComObject(mailItem);
                if (outlookApp != null) Marshal.FinalReleaseComObject(outlookApp);
            }
        }
    }
}