using DocumentFormat.OpenXml.Packaging;
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

            // Set default dates
            DatePickup.SelectedDate = DateTime.Today;
            DateDeliver.SelectedDate = DateTime.Today.AddDays(1);
            CapacDataDatePicker.SelectedDate = DateTime.Today;

            // Handle text box validation
            var textBoxes = new[] { nrTank, DescriptionTextBox, Address1TextBox, Address2TextBox, MaxDaysTextBox, CapacClientTextBox, CapacRutaTextBox, CapacNumarInmatriculareTextBox, CapacTransportatorTextBox, CapacPretTextBox, CapacCurrencyTextBox, CapacCantitateTextBox, CapacFacturaClientTextBox, CapacFacturaCarausTextBox };
            foreach (var textBox in textBoxes)
            {
                textBox.TextChanged += Input_TextChanged;
            }

            // Handle date picker validation
            DatePickup.SelectedDateChanged += DatePicker_SelectedDateChanged;
            DateDeliver.SelectedDateChanged += DatePicker_SelectedDateChanged;
        }

        private void Input_TextChanged(object sender, TextChangedEventArgs e)
        {
            TextBox textBox = (TextBox)sender;
            if (!string.IsNullOrWhiteSpace(textBox.Text))
            {
                textBox.BorderBrush = new SolidColorBrush((System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#20C997"));
            }
            else
            {
                textBox.BorderBrush = new SolidColorBrush((System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#007BFF"));
            }
        }

        private void DatePicker_SelectedDateChanged(object sender, SelectionChangedEventArgs e)
        {
            DatePicker datePicker = (DatePicker)sender;

            // Simple validation - just check if a date is selected
            if (datePicker.SelectedDate.HasValue)
            {
                datePicker.BorderBrush = new SolidColorBrush((System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#20C997"));
            }
            else
            {
                datePicker.BorderBrush = new SolidColorBrush((System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#007BFF"));
            }
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
            string datePickupSlash = DatePickup.SelectedDate.HasValue ? DatePickup.SelectedDate.Value.ToString("dd/MM/yyyy") : string.Empty;
            string dateDeliverSlash = DateDeliver.SelectedDate.HasValue ? DateDeliver.SelectedDate.Value.ToString("dd/MM/yyyy") : string.Empty;

            // Template expects commas as separators per example: 21,11,2023
            string datePickup = DatePickup.SelectedDate.HasValue ? DatePickup.SelectedDate.Value.ToString("dd,MM,yyyy") : string.Empty;
            string dateDeliver = DateDeliver.SelectedDate.HasValue ? DateDeliver.SelectedDate.Value.ToString("dd,MM,yyyy") : string.Empty;

            string dataCapac = CapacDataDatePicker.SelectedDate.HasValue ? CapacDataDatePicker.SelectedDate.Value.ToString("dd/MM/yyyy") : string.Empty;

            string Get(string? s) => s?.Trim() ?? string.Empty;
            string qty = Get(CapacCantitateTextBox.Text);
            string qtyWithUnit = string.IsNullOrEmpty(qty) ? string.Empty : $"{qty} KG";

            var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                // Comanda Transport placeholders
                { "nr. Tank", nrTank.Text?.Trim() ?? string.Empty },
                { "21,11,2023", datePickup },
                { "24,11,2023", dateDeliver },
                { "Adresa de incarcare", Address1TextBox.Text?.Trim() ?? string.Empty },
                { "Adresa de descarcare", Address2TextBox.Text?.Trim() ?? string.Empty },
                { "Descriere marfa:", DescriptionTextBox.Text?.Trim() ?? string.Empty},
                { "PREŢ NEGOCIAT:", $"PREŢ NEGOCIAT: {BuildPrice()}" },
                { "maxim 45 zile", BuildMaxDays() },
                
                // CAPAC placeholders
                { "CLIENT:", $"CLIENT: {Get(CapacClientTextBox.Text)}" },
                { "RUTA:", $"RUTA: {Get(CapacRutaTextBox.Text)}" },
                { "DATA:", $"DATA: {dataCapac}" },
                { "NUMAR INMATRICULARE:", $"NUMAR INMATRICULARE: {Get(CapacNumarInmatriculareTextBox.Text)}" },
                { "TRANSPORTATOR:", $"TRANSPORTATOR: {Get(CapacTransportatorTextBox.Text)}" },
                { "PRET:", $"PRET: {BuildPrice()}" },
                { "Cantitate incarcata:", $"Cantitate incarcata: {qtyWithUnit}" },
                { "Factura client:", $"Factura client: {Get(CapacFacturaClientTextBox.Text)}" },
                { "Factura caraus:", $"Factura caraus: {Get(CapacFacturaCarausTextBox.Text)}" },
                
                // Additional fallback general tokens
                { "{{DatePickup}}", datePickupSlash },
                { "{{DateDeliver}}", dateDeliverSlash },
                { "{{Today}}", DateTime.Now.ToString("dd/MM/yyyy") },
                { "{{NrTank}}", nrTank.Text?.Trim() ?? string.Empty },
                { "{{Description}}", DescriptionTextBox.Text?.Trim() ?? string.Empty },
                { "{{Address1}}", Address1TextBox.Text?.Trim() ?? string.Empty },
                { "{{Address2}}", Address2TextBox.Text?.Trim() ?? string.Empty },
                { "{{Price}}", BuildPrice() },
                { "{{MaxDays}}", BuildMaxDays() }
            };

            return map;
        }

        private string BuildPrice()
        {
            string price = CapacPretTextBox.Text?.Trim() ?? string.Empty;
            string currency = CapacCurrencyTextBox.Text?.Trim() ?? string.Empty;
            string combined = (price + " " + currency).Trim();
            return string.IsNullOrEmpty(combined) ? string.Empty : combined;
        }

        private string BuildMaxDays()
        {
            string maxDays = MaxDaysTextBox.Text?.Trim() ?? string.Empty;
            return string.IsNullOrEmpty(maxDays) ? string.Empty : $"maxim {maxDays} zile";
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
            // Capture original paragraph text before any run edits
            string originalParagraphText = paragraph.Text;

            // First try run-level replacements (preserves most formatting when placeholders are not split)
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

            // If placeholders are split across runs, fall back to paragraph-level rebuild
            // Only do this if no run-level change occurred to avoid double application
            if (!anyRunChanged)
            {
                string newParaText = ReplaceAll(originalParagraphText, replacements);
                if (!string.Equals(originalParagraphText, newParaText, StringComparison.Ordinal))
                {
                    // Remove all runs and set a single run with replaced text
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
                return false; // Outlook not installed
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