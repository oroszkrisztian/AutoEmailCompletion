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
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;
using EmailCompleteApp.Services;
using EmailCompleteApp.Models;

namespace EmailCompleteApp.Pages
{
    public partial class ComandaTransport : UserControl
    {
        private readonly SearchService _searchService;
        private DispatcherTimer _searchTimer;
        private DispatcherTimer _transportatorSearchTimer;
        private DispatcherTimer _incarcareSearchTimer;
        private DispatcherTimer _descarcareSearchTimer;
        private bool _isUpdatingComboBox = false;
        private bool _isUpdatingTransportatorComboBox = false;
        private bool _isUpdatingIncarcareComboBox = false;
        private bool _isUpdatingDescarcareComboBox = false;

        public ComandaTransport()
        {
            InitializeComponent();
            
            _searchService = new SearchService();
            
            // Initialize search timers for debouncing
            _searchTimer = new DispatcherTimer
            {
                Interval = TimeSpan.FromMilliseconds(100)
            };
            _searchTimer.Tick += OnSearchTimerTick;

            _transportatorSearchTimer = new DispatcherTimer
            {
                Interval = TimeSpan.FromMilliseconds(100)
            };
            _transportatorSearchTimer.Tick += OnTransportatorSearchTimerTick;

            _incarcareSearchTimer = new DispatcherTimer
            {
                Interval = TimeSpan.FromMilliseconds(100)
            };
            _incarcareSearchTimer.Tick += OnIncarcareSearchTimerTick;

            _descarcareSearchTimer = new DispatcherTimer
            {
                Interval = TimeSpan.FromMilliseconds(100)
            };
            _descarcareSearchTimer.Tick += OnDescarcareSearchTimerTick;

            MonedaComboBox.SelectedIndex = 0;
            TipComboBox.SelectedIndex = 0;
            if (TransportatorMonedaComboBox != null) TransportatorMonedaComboBox.SelectedIndex = 0;
            if (TransportatorTipComboBox != null) TransportatorTipComboBox.SelectedIndex = 0;
            if (TipAdrComboBox != null) TipAdrComboBox.SelectedIndex = 0;

            // Set default dates (if present)
            if (DataIncarcareDatePicker != null) DataIncarcareDatePicker.SelectedDate = DateTime.Today;
            if (DataDescarcareDatePicker != null) DataDescarcareDatePicker.SelectedDate = DateTime.Today.AddDays(1);
        }

        private void OnClientTextChanged(object sender, TextChangedEventArgs e)
        {
            if (_isUpdatingComboBox) return;

            // Stop and restart the timer to debounce the search
            _searchTimer.Stop();
            _searchTimer.Start();
        }

        private async void OnSearchTimerTick(object sender, EventArgs e)
        {
            _searchTimer.Stop();
            
            var searchText = ClientComboBox.Text?.Trim();
            
            // Don't clear items if no text - just close dropdown
            if (string.IsNullOrEmpty(searchText))
            {
                ClientComboBox.IsDropDownOpen = false;
                return;
            }
            
            try
            {
                var clientNames = await _searchService.SearchClientNamesAsync(searchText);
                
                _isUpdatingComboBox = true;
                
                // Store current text and selection
                var currentText = ClientComboBox.Text;
                var currentSelectionStart = ClientComboBox.IsEditable ? GetTextBoxFromComboBox(ClientComboBox)?.SelectionStart ?? 0 : 0;
                
                // Only update items if we have results
                if (clientNames.Any())
                {
                    ClientComboBox.Items.Clear();
                    foreach (var name in clientNames)
                    {
                        ClientComboBox.Items.Add(name);
                    }
                    
                    // Show dropdown with results
                    ClientComboBox.IsDropDownOpen = true;
                }
                else
                {
                    // Close dropdown if no results, but don't clear existing items
                    ClientComboBox.IsDropDownOpen = false;
                }
                
                // Always restore text and selection
                ClientComboBox.Text = currentText;
                if (ClientComboBox.IsEditable)
                {
                    var textBox = GetTextBoxFromComboBox(ClientComboBox);
                    if (textBox != null)
                    {
                        textBox.SelectionStart = currentSelectionStart;
                        textBox.SelectionLength = 0;
                    }
                }
                
                _isUpdatingComboBox = false;
            }
            catch
            {
                _isUpdatingComboBox = false;
            }
        }

        private void OnClientSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_isUpdatingComboBox) return;
            
            // When user selects a client from dropdown, close it and remove focus
            if (ClientComboBox.SelectedItem != null)
            {
                _isUpdatingComboBox = true; // Prevent text change events during selection
                
                // Stop the search timer to prevent reopening
                _searchTimer.Stop();
                
                // Close dropdown immediately
                ClientComboBox.IsDropDownOpen = false;
                
                // Use dispatcher to delay the focus change and flag reset
                Dispatcher.BeginInvoke(new Action(async () =>
                {
                    await Task.Delay(50); // Small delay to ensure dropdown is closed
                    
                    // Clear text selection to show plain text
                    var textBox = GetTextBoxFromComboBox(ClientComboBox);
                    if (textBox != null)
                    {
                        textBox.SelectionStart = textBox.Text.Length; // Move cursor to end
                        textBox.SelectionLength = 0; // Clear selection
                    }
                    
                    // Remove focus from the ComboBox to prevent reopening
                    ClientComboBox.MoveFocus(new TraversalRequest(FocusNavigationDirection.Next));
                    
                    await Task.Delay(100); // Additional delay before resetting flag
                    _isUpdatingComboBox = false;
                }), System.Windows.Threading.DispatcherPriority.Background);
            }
        }

        private static TextBox? GetTextBoxFromComboBox(ComboBox comboBox)
        {
            return comboBox.Template?.FindName("PART_EditableTextBox", comboBox) as TextBox;
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

                string mergedTemplatePath = Path.Combine(docDir, "comanda.docx");

                string generatedDir = Path.Combine(docDir, "Generated");
                Directory.CreateDirectory(generatedDir);

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
                { "Client", ClientComboBox.Text?.Trim() ?? string.Empty },
                { "Tarif", TarifTextBox.Text?.Trim() ?? string.Empty },
                { "Primit", PrimitTextBox.Text?.Trim() ?? string.Empty },
                { "Moneda", (MonedaComboBox.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? string.Empty },
                { "Tip", (TipComboBox.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? string.Empty },
                { "Transportator", TransportatorComboBox.Text?.Trim() ?? string.Empty },
                { "Transportator Tarif", TransportatorTarifTextBox.Text?.Trim() ?? string.Empty },
                { "Oferit", OferitTextBox.Text?.Trim() ?? string.Empty },
                { "Transportator Moneda", (TransportatorMonedaComboBox.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? string.Empty },
                { "Transportator Tip", (TransportatorTipComboBox.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? string.Empty },
                { "Data Incarcare", DataIncarcareDatePicker.SelectedDate?.ToString("dd/MM/yyyy") ?? string.Empty },
                { "Data Descarcare", DataDescarcareDatePicker.SelectedDate?.ToString("dd/MM/yyyy") ?? string.Empty },
                { "Produs", ProdusTextBox.Text?.Trim() ?? string.Empty },
                { "Cantitate", CantitateTextBox.Text?.Trim() ?? string.Empty },
                { "Tip ADR", (TipAdrComboBox.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? string.Empty },
                { "Clasa", ClasaTextBox.Text?.Trim() ?? string.Empty },
                { "UM", UMTextBox.Text?.Trim() ?? string.Empty },
                { "Numar Inmatriculare", NumarInmatriculareTextBox.Text?.Trim() ?? string.Empty },
                { "Adresa Incarcare", LocatieIncarcareComboBox.Text?.Trim() ?? string.Empty },
                { "Adresa Descarcare", LocatieDescarcareComboBox.Text?.Trim() ?? string.Empty },
                { "Termen Plata", MaxDaysTextBox.Text?.Trim() ?? string.Empty }
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

        private void OnTransportatorTextChanged(object sender, TextChangedEventArgs e)
        {
            if (_isUpdatingTransportatorComboBox) return;

            // Stop and restart the timer to debounce the search
            _transportatorSearchTimer.Stop();
            _transportatorSearchTimer.Start();
        }

        private async void OnTransportatorSearchTimerTick(object sender, EventArgs e)
        {
            _transportatorSearchTimer.Stop();
            
            var searchText = TransportatorComboBox.Text?.Trim();
            
            // Don't clear items if no text - just close dropdown
            if (string.IsNullOrEmpty(searchText))
            {
                TransportatorComboBox.IsDropDownOpen = false;
                return;
            }
            
            try
            {
                var transportatorNames = await _searchService.SearchTransportatorNamesAsync(searchText);
                
                _isUpdatingTransportatorComboBox = true;
                
                // Store current text and selection
                var currentText = TransportatorComboBox.Text;
                var currentSelectionStart = TransportatorComboBox.IsEditable ? GetTextBoxFromComboBox(TransportatorComboBox)?.SelectionStart ?? 0 : 0;
                
                // Only update items if we have results
                if (transportatorNames.Any())
                {
                    TransportatorComboBox.Items.Clear();
                    foreach (var name in transportatorNames)
                    {
                        TransportatorComboBox.Items.Add(name);
                    }
                    
                    // Show dropdown with results
                    TransportatorComboBox.IsDropDownOpen = true;
                }
                else
                {
                    // Close dropdown if no results, but don't clear existing items
                    TransportatorComboBox.IsDropDownOpen = false;
                }
                
                // Always restore text and selection
                TransportatorComboBox.Text = currentText;
                if (TransportatorComboBox.IsEditable)
                {
                    var textBox = GetTextBoxFromComboBox(TransportatorComboBox);
                    if (textBox != null)
                    {
                        textBox.SelectionStart = currentSelectionStart;
                        textBox.SelectionLength = 0;
                    }
                }
                
                _isUpdatingTransportatorComboBox = false;
            }
            catch
            {
                _isUpdatingTransportatorComboBox = false;
            }
        }

        private void OnTransportatorSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_isUpdatingTransportatorComboBox) return;
            
            // When user selects a transportator from dropdown, close it and remove focus
            if (TransportatorComboBox.SelectedItem != null)
            {
                _isUpdatingTransportatorComboBox = true; // Prevent text change events during selection
                
                // Stop the search timer to prevent reopening
                _transportatorSearchTimer.Stop();
                
                // Close dropdown immediately
                TransportatorComboBox.IsDropDownOpen = false;
                
                // Use dispatcher to delay the focus change and flag reset
                Dispatcher.BeginInvoke(new Action(async () =>
                {
                    await Task.Delay(50); // Small delay to ensure dropdown is closed
                    
                    // Clear text selection to show plain text
                    var textBox = GetTextBoxFromComboBox(TransportatorComboBox);
                    if (textBox != null)
                    {
                        textBox.SelectionStart = textBox.Text.Length; // Move cursor to end
                        textBox.SelectionLength = 0; // Clear selection
                    }
                    
                    // Remove focus from the ComboBox to prevent reopening
                    TransportatorComboBox.MoveFocus(new TraversalRequest(FocusNavigationDirection.Next));
                    
                    await Task.Delay(100); // Additional delay before resetting flag
                    _isUpdatingTransportatorComboBox = false;
                }), System.Windows.Threading.DispatcherPriority.Background);
            }
        }

        private void OnIncarcareTextChanged(object sender, TextChangedEventArgs e)
        {
            if (_isUpdatingIncarcareComboBox) return;

            // Stop and restart the timer to debounce the search
            _incarcareSearchTimer.Stop();
            _incarcareSearchTimer.Start();
        }

        private async void OnIncarcareSearchTimerTick(object sender, EventArgs e)
        {
            _incarcareSearchTimer.Stop();
            
            var searchText = LocatieIncarcareComboBox.Text?.Trim();
            
            // Don't clear items if no text - just close dropdown
            if (string.IsNullOrEmpty(searchText))
            {
                LocatieIncarcareComboBox.IsDropDownOpen = false;
                return;
            }
            
            try
            {
                var addresses = await _searchService.SearchLocationAddressesAsync(searchText);
                
                _isUpdatingIncarcareComboBox = true;
                
                // Store current text and selection
                var currentText = LocatieIncarcareComboBox.Text;
                var currentSelectionStart = LocatieIncarcareComboBox.IsEditable ? GetTextBoxFromComboBox(LocatieIncarcareComboBox)?.SelectionStart ?? 0 : 0;
                
                // Only update items if we have results
                if (addresses.Any())
                {
                    LocatieIncarcareComboBox.Items.Clear();
                    foreach (var address in addresses)
                    {
                        LocatieIncarcareComboBox.Items.Add(address);
                    }
                    
                    // Show dropdown with results
                    LocatieIncarcareComboBox.IsDropDownOpen = true;
                }
                else
                {
                    // Close dropdown if no results, but don't clear existing items
                    LocatieIncarcareComboBox.IsDropDownOpen = false;
                }
                
                // Always restore text and selection
                LocatieIncarcareComboBox.Text = currentText;
                if (LocatieIncarcareComboBox.IsEditable)
                {
                    var textBox = GetTextBoxFromComboBox(LocatieIncarcareComboBox);
                    if (textBox != null)
                    {
                        textBox.SelectionStart = currentSelectionStart;
                        textBox.SelectionLength = 0;
                    }
                }
                
                _isUpdatingIncarcareComboBox = false;
            }
            catch
            {
                _isUpdatingIncarcareComboBox = false;
            }
        }

        private void OnIncarcareSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_isUpdatingIncarcareComboBox) return;
            
            // When user selects an address from dropdown, close it and remove focus
            if (LocatieIncarcareComboBox.SelectedItem != null)
            {
                _isUpdatingIncarcareComboBox = true; // Prevent text change events during selection
                
                // Stop the search timer to prevent reopening
                _incarcareSearchTimer.Stop();
                
                // Close dropdown immediately
                LocatieIncarcareComboBox.IsDropDownOpen = false;
                
                // Use dispatcher to delay the focus change and flag reset
                Dispatcher.BeginInvoke(new Action(async () =>
                {
                    await Task.Delay(50); // Small delay to ensure dropdown is closed
                    
                    // Clear text selection to show plain text
                    var textBox = GetTextBoxFromComboBox(LocatieIncarcareComboBox);
                    if (textBox != null)
                    {
                        textBox.SelectionStart = textBox.Text.Length; // Move cursor to end
                        textBox.SelectionLength = 0; // Clear selection
                    }
                    
                    // Remove focus from the ComboBox to prevent reopening
                    LocatieIncarcareComboBox.MoveFocus(new TraversalRequest(FocusNavigationDirection.Next));
                    
                    await Task.Delay(100); // Additional delay before resetting flag
                    _isUpdatingIncarcareComboBox = false;
                }), System.Windows.Threading.DispatcherPriority.Background);
            }
        }

        private void OnDescarcareTextChanged(object sender, TextChangedEventArgs e)
        {
            if (_isUpdatingDescarcareComboBox) return;

            // Stop and restart the timer to debounce the search
            _descarcareSearchTimer.Stop();
            _descarcareSearchTimer.Start();
        }

        private async void OnDescarcareSearchTimerTick(object sender, EventArgs e)
        {
            _descarcareSearchTimer.Stop();
            
            var searchText = LocatieDescarcareComboBox.Text?.Trim();
            
            // Don't clear items if no text - just close dropdown
            if (string.IsNullOrEmpty(searchText))
            {
                LocatieDescarcareComboBox.IsDropDownOpen = false;
                return;
            }
            
            try
            {
                var addresses = await _searchService.SearchLocationAddressesAsync(searchText);
                
                _isUpdatingDescarcareComboBox = true;
                
                // Store current text and selection
                var currentText = LocatieDescarcareComboBox.Text;
                var currentSelectionStart = LocatieDescarcareComboBox.IsEditable ? GetTextBoxFromComboBox(LocatieDescarcareComboBox)?.SelectionStart ?? 0 : 0;
                
                // Only update items if we have results
                if (addresses.Any())
                {
                    LocatieDescarcareComboBox.Items.Clear();
                    foreach (var address in addresses)
                    {
                        LocatieDescarcareComboBox.Items.Add(address);
                    }
                    
                    // Show dropdown with results
                    LocatieDescarcareComboBox.IsDropDownOpen = true;
                }
                else
                {
                    // Close dropdown if no results, but don't clear existing items
                    LocatieDescarcareComboBox.IsDropDownOpen = false;
                }
                
                // Always restore text and selection
                LocatieDescarcareComboBox.Text = currentText;
                if (LocatieDescarcareComboBox.IsEditable)
                {
                    var textBox = GetTextBoxFromComboBox(LocatieDescarcareComboBox);
                    if (textBox != null)
                    {
                        textBox.SelectionStart = currentSelectionStart;
                        textBox.SelectionLength = 0;
                    }
                }
                
                _isUpdatingDescarcareComboBox = false;
            }
            catch
            {
                _isUpdatingDescarcareComboBox = false;
            }
        }

        private void OnDescarcareSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_isUpdatingDescarcareComboBox) return;
            
            // When user selects an address from dropdown, close it and remove focus
            if (LocatieDescarcareComboBox.SelectedItem != null)
            {
                _isUpdatingDescarcareComboBox = true; // Prevent text change events during selection
                
                // Stop the search timer to prevent reopening
                _descarcareSearchTimer.Stop();
                
                // Close dropdown immediately
                LocatieDescarcareComboBox.IsDropDownOpen = false;
                
                // Use dispatcher to delay the focus change and flag reset
                Dispatcher.BeginInvoke(new Action(async () =>
                {
                    await Task.Delay(50); // Small delay to ensure dropdown is closed
                    
                    // Clear text selection to show plain text
                    var textBox = GetTextBoxFromComboBox(LocatieDescarcareComboBox);
                    if (textBox != null)
                    {
                        textBox.SelectionStart = textBox.Text.Length; // Move cursor to end
                        textBox.SelectionLength = 0; // Clear selection
                    }
                    
                    // Remove focus from the ComboBox to prevent reopening
                    LocatieDescarcareComboBox.MoveFocus(new TraversalRequest(FocusNavigationDirection.Next));
                    
                    await Task.Delay(100); // Additional delay before resetting flag
                    _isUpdatingDescarcareComboBox = false;
                }), System.Windows.Threading.DispatcherPriority.Background);
            }
        }
    }
}