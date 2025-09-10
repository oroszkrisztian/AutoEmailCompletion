using System;
using System.Collections.Generic;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using NPOI.XWPF.UserModel;
using System.Diagnostics;
using System.Runtime.InteropServices;
using Word = Microsoft.Office.Interop.Word;
using DocumentFormat.OpenXml.Packaging;
using OpenXmlPowerTools;

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
                textBox.BorderBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#20C997"));
            }
            else
            {
                textBox.BorderBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#007BFF"));
            }
        }

        private void DatePicker_SelectedDateChanged(object sender, SelectionChangedEventArgs e)
        {
            DatePicker datePicker = (DatePicker)sender;

            // Simple validation - just check if a date is selected
            if (datePicker.SelectedDate.HasValue)
            {
                datePicker.BorderBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#20C997"));
            }
            else
            {
                datePicker.BorderBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#007BFF"));
            }
        }

        private void OnSendClick(object sender, System.Windows.RoutedEventArgs e)
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
                string comandaTemplatePath = Path.Combine(docDir, "Comanda_transport.docx");
                string capacTemplatePath = Path.Combine(docDir, "CAPAC_dosar.docx");

                string generatedDir = Path.Combine(docDir, "Generated");
                Directory.CreateDirectory(generatedDir);

                // Human-readable, Windows-safe timestamp (no colons)
                string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH.mm.ss");
                var comandaReplacements = BuildComandaReplacements();
                var capacReplacements = BuildCapacReplacements();

                string comandaOutputPath = Path.Combine(generatedDir, $"Comanda transport - {timestamp}.docx");
                string capacOutputPath = Path.Combine(generatedDir, $"CAPAC dosar - {timestamp}.docx");

                // Check templates exist
                if (!File.Exists(comandaTemplatePath))
                {
                    MessageBox.Show($"No template found. Add 'Comanda_transport.docx' under: {docDir}",
                                  "Template Missing", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }
                if (!File.Exists(capacTemplatePath))
                {
                    MessageBox.Show($"No template found. Add 'CAPAC_dosar.docx' under: {docDir}",
                                  "Template Missing", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                // Generate both documents
                GenerateWordDocumentFromTemplate(comandaTemplatePath, comandaOutputPath, comandaReplacements);
                GenerateWordDocumentFromTemplate(capacTemplatePath, capacOutputPath, capacReplacements);

                // Merge into single DOCX (preserving formatting) without requiring Word
                string mergedOutputPath = Path.Combine(generatedDir, $"CAPAC+Comanda transport - {timestamp}.docx");
                try
                {
                    MergeDocxWithOpenXmlPowerTools(capacOutputPath, comandaOutputPath, mergedOutputPath);
                }
                catch (Exception mEx)
                {
                    Debug.WriteLine($"Merging with Word failed: {mEx.ToString()}");
                    MessageBox.Show($"Generated files, but failed to merge automatically.\n\nError: {mEx.Message}\n\n" +
                                    $"CAPAC: {capacOutputPath}\nComanda: {comandaOutputPath}",
                                    "Partial Success", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                MessageBox.Show($"Generated and merged successfully.\n\nMerged file: {mergedOutputPath}",
                                "Success", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to generate document.\n\nError: {ex.Message}",
                              "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private Dictionary<string, string> BuildComandaReplacements()
        {
            string datePickupSlash = DatePickup.SelectedDate.HasValue ? DatePickup.SelectedDate.Value.ToString("dd/MM/yyyy") : string.Empty;
            string dateDeliverSlash = DateDeliver.SelectedDate.HasValue ? DateDeliver.SelectedDate.Value.ToString("dd/MM/yyyy") : string.Empty;

            // Template expects commas as separators per example: 21,11,2023
            string datePickup = DatePickup.SelectedDate.HasValue ? DatePickup.SelectedDate.Value.ToString("dd,MM,yyyy") : string.Empty;
            string dateDeliver = DateDeliver.SelectedDate.HasValue ? DateDeliver.SelectedDate.Value.ToString("dd,MM,yyyy") : string.Empty;

            var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                // Exact placeholders from the template document
                { "nr. Tank", nrTank.Text?.Trim() ?? string.Empty },
                { "21,11,2023", datePickup },
                { "24,11,2023", dateDeliver },
                { "Adresa de incarcare", Address1TextBox.Text?.Trim() ?? string.Empty },
                { "Adresa de descarcare", Address2TextBox.Text?.Trim() ?? string.Empty },
                { "Descriere marfa:", DescriptionTextBox.Text?.Trim() ?? string.Empty},
                { "PREŢ NEGOCIAT:", $"PREŢ NEGOCIAT: {BuildPrice()}" },
                { "maxim 45 zile", BuildMaxDays() },
                
                // Additional fallback general tokens if they also exist in the doc
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

        private Dictionary<string, string> BuildCapacReplacements()
        {
            string dataCapac = CapacDataDatePicker.SelectedDate.HasValue ? CapacDataDatePicker.SelectedDate.Value.ToString("dd/MM/yyyy") : string.Empty;

            string Get(string? s) => s?.Trim() ?? string.Empty;
            string qty = Get(CapacCantitateTextBox.Text);
            string qtyWithUnit = string.IsNullOrEmpty(qty) ? string.Empty : $"{qty} KG";

            var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "CLIENT:", $"CLIENT: {Get(CapacClientTextBox.Text)}" },
                { "RUTA:", $"RUTA: {Get(CapacRutaTextBox.Text)}" },
                { "DATA:", $"DATA: {dataCapac}" },
                { "NUMAR INMATRICULARE:", $"NUMAR INMATRICULARE: {Get(CapacNumarInmatriculareTextBox.Text)}" },
                { "TRANSPORTATOR:", $"TRANSPORTATOR: {Get(CapacTransportatorTextBox.Text)}" },
                { "PRET:", $"PRET: {BuildPrice()}" },
                { "Cantitate incarcata:", $"Cantitate incarcata: {qtyWithUnit}" },
                { "Factura client:", $"Factura client: {Get(CapacFacturaClientTextBox.Text)}" },
                { "Factura caraus:", $"Factura caraus: {Get(CapacFacturaCarausTextBox.Text)}" }
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

        private static void MergeDocxWithWord(string firstDocPath, string secondDocPath, string mergedOutputPath)
        {
            if (!File.Exists(firstDocPath)) throw new FileNotFoundException("First document not found", firstDocPath);
            if (!File.Exists(secondDocPath)) throw new FileNotFoundException("Second document not found", secondDocPath);

            Word.Application wordApp = null;
            Word.Document mergedDoc = null;
            object missing = Type.Missing;

            try
            {
                // Check if Word is installed
                try
                {
                    wordApp = new Word.Application();
                }
                catch (COMException ex) when (ex.HResult == unchecked((int)0x80040154))
                {
                    throw new InvalidOperationException("Microsoft Word is not installed. Cannot merge documents.");
                }

                wordApp.Visible = false;
                wordApp.ScreenUpdating = false;
                wordApp.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;

                // Create a new document
                mergedDoc = wordApp.Documents.Add(ref missing, ref missing, ref missing, ref missing);

                // Insert first document
                Word.Range range = mergedDoc.Content;
                range.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                range.InsertFile(firstDocPath, ref missing, ref missing, ref missing, ref missing);

                // Insert page break
                range.InsertBreak(Word.WdBreakType.wdPageBreak);

                // Insert second document
                range.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                range.InsertFile(secondDocPath, ref missing, ref missing, ref missing, ref missing);

                // Save merged document
                mergedDoc.SaveAs2(mergedOutputPath, FileFormat: Word.WdSaveFormat.wdFormatXMLDocument);
            }
            finally
            {
                // Cleanup COM objects
                if (mergedDoc != null)
                {
                    mergedDoc.Close(ref missing, ref missing, ref missing);
                    Marshal.ReleaseComObject(mergedDoc);
                }
                if (wordApp != null)
                {
                    wordApp.Quit(ref missing, ref missing, ref missing);
                    Marshal.ReleaseComObject(wordApp);
                }

                // Force garbage collection to clean up remaining COM references
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        private static void MergeDocxWithOpenXmlPowerTools(string firstDocPath, string secondDocPath, string mergedOutputPath)
        {
            if (!File.Exists(firstDocPath)) throw new FileNotFoundException("First document not found", firstDocPath);
            if (!File.Exists(secondDocPath)) throw new FileNotFoundException("Second document not found", secondDocPath);

            var sources = new List<Source>()
            {
                new Source(new WmlDocument(firstDocPath), false),
                new Source(new WmlDocument(secondDocPath), true) 
            };

            var merged = DocumentBuilder.BuildDocument(sources);
            merged.SaveAs(mergedOutputPath);
        }

        private static string ConvertDocxToPdfWithWord(string docxPath)
        {
            if (!File.Exists(docxPath))
                throw new FileNotFoundException("DOCX not found", docxPath);

            string pdfPath = Path.ChangeExtension(docxPath, ".pdf");

            Word.Application? wordApp = null;
            Word.Document? doc = null;
            try
            {
                wordApp = new Word.Application { Visible = false, ScreenUpdating = false, DisplayAlerts = Word.WdAlertLevel.wdAlertsNone };
                object readOnly = true;
                object isVisible = false;
                object missing = Type.Missing;
                object fileName = docxPath;
                doc = wordApp.Documents.Open(ref fileName, ReadOnly: ref readOnly, Visible: ref isVisible);

                object outputFileName = pdfPath;
                var exportFormat = Word.WdExportFormat.wdExportFormatPDF;
                doc.ExportAsFixedFormat(OutputFileName: pdfPath,
                                        ExportFormat: exportFormat,
                                        OpenAfterExport: false,
                                        OptimizeFor: Word.WdExportOptimizeFor.wdExportOptimizeForPrint,
                                        Range: Word.WdExportRange.wdExportAllDocument,
                                        From: 0,
                                        To: 0,
                                        Item: Word.WdExportItem.wdExportDocumentContent,
                                        IncludeDocProps: true,
                                        KeepIRM: true,
                                        CreateBookmarks: Word.WdExportCreateBookmarks.wdExportCreateNoBookmarks,
                                        DocStructureTags: true,
                                        BitmapMissingFonts: true,
                                        UseISO19005_1: false);

                return pdfPath;
            }
            finally
            {
                if (doc != null)
                {
                    try { doc.Close(SaveChanges: false); } catch { }
                    Marshal.FinalReleaseComObject(doc);
                }
                if (wordApp != null)
                {
                    try { wordApp.Quit(SaveChanges: false); } catch { }
                    Marshal.FinalReleaseComObject(wordApp);
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
                mailType.GetProperty("Body")?.SetValue(mailItem, "Va rugam gasiti atasat documentul in format PDF.");

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