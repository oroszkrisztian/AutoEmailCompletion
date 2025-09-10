using System;
using System.Collections.Generic;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using NPOI.XWPF.UserModel;

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

            // Handle text box validation
            var textBoxes = new[] { nrTank, DescriptionTextBox, Address1TextBox, Address2TextBox, PriceTextBox, CurrencyTextBox, MaxDaysTextBox };
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
                string templatePath = Path.Combine(docDir, "Comanda_transport.docx");

                string generatedDir = Path.Combine(docDir, "Generated");
                Directory.CreateDirectory(generatedDir);

                string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                var replacements = BuildReplacements();

                string outputPath = Path.Combine(generatedDir, $"Comanda_transport_{timestamp}.docx");

                // Check if template exists
                if (!File.Exists(templatePath))
                {
                    MessageBox.Show($"No template found. Add 'Comanda_transport.docx' under: {docDir}",
                                  "Template Missing", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                // Generate document using NPOI
                GenerateWordDocumentFromTemplate(templatePath, outputPath, replacements);

                MessageBox.Show($"Document generated successfully!\n\nSaved to: {outputPath}",
                              "Success", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to generate document.\n\nError: {ex.Message}",
                              "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private Dictionary<string, string> BuildReplacements()
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
                { "Descriere marfa:", DescriptionTextBox.Text?.Trim() ?? string.Empty },
                { "PREŢ NEGOCIAT", BuildPrice() },
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

        private string BuildPrice()
        {
            string price = PriceTextBox.Text?.Trim() ?? string.Empty;
            string currency = CurrencyTextBox.Text?.Trim() ?? string.Empty;
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
            // First try run-level replacements (preserves most formatting when placeholders are not split)
            var runs = paragraph.Runs;
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
                    }
                }
            }

            // If placeholders are split across runs, fall back to paragraph-level rebuild
            string paraText = paragraph.Text;
            string newParaText = ReplaceAll(paraText, replacements);
            if (!string.Equals(paraText, newParaText, StringComparison.Ordinal))
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
    }
}