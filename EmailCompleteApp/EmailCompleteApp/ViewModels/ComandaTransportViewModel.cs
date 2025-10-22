using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using NPOI.XWPF.UserModel;
using EmailCompleteApp.Services;
using EmailCompleteApp.Models;
using System.CodeDom;

namespace EmailCompleteApp.ViewModels
{
    public partial class ComandaTransportViewModel : ObservableObject
    {
        private readonly ClientsViewModel _clientViewmodel;
        private readonly SearchService _searchService;
        private CancellationTokenSource? _clientSearchCts;
        private CancellationTokenSource? _transportatorSearchCts;
        private CancellationTokenSource? _incarcareSearchCts;
        private CancellationTokenSource? _descarcareSearchCts;

        [ObservableProperty]
        private string _numarComanda = string.Empty;
        
        [ObservableProperty]
        private string _numarClient = string.Empty;

        [ObservableProperty]
        private string _client = string.Empty;

        [ObservableProperty]
        private string _tarif = string.Empty;

        [ObservableProperty]
        private int _monedaIndex = 0;

        [ObservableProperty]
        private int _tipIndex = 0;

        [ObservableProperty]
        private string _transportator = string.Empty;

        [ObservableProperty]
        private string _transportatorTarif = string.Empty;

        [ObservableProperty]
        private int _transportatorMonedaIndex = 0;

        [ObservableProperty]
        private int _transportatorTipIndex = 0;

        [ObservableProperty]
        private DateTime? _dataIncarcare = DateTime.Today;

        [ObservableProperty]
        private DateTime? _dataDescarcare = DateTime.Today.AddDays(1);

        [ObservableProperty]
        private string _produs = string.Empty;

        [ObservableProperty]
        private string _cantitate = string.Empty;

        [ObservableProperty]
        private int _tipAdrIndex = 0;

        [ObservableProperty]
        private string _clasa = string.Empty;

        [ObservableProperty]
        private string _un = string.Empty;

        [ObservableProperty]
        private string _numarInmatriculare = string.Empty;

        [ObservableProperty]
        private string _locatieIncarcare = string.Empty;

        [ObservableProperty]
        private string _locatieDescarcare = string.Empty;

        [ObservableProperty]
        private string _termenPlata = string.Empty;

        [ObservableProperty]
        private bool _isClientDropDownOpen = false;

        [ObservableProperty]
        private bool _isTransportatorDropDownOpen = false;

        [ObservableProperty]
        private bool _isIncarcareDropDownOpen = false;

        [ObservableProperty]
        private bool _isDescarcareDropDownOpen = false;

        [ObservableProperty]
        private bool _isSendingEmail = false;

        public ObservableCollection<string> ClientSuggestions { get; } = new();
        public ObservableCollection<string> TransportatorSuggestions { get; } = new();
        public ObservableCollection<string> IncarcareSuggestions { get; } = new();
        public ObservableCollection<string> DescarcareSuggestions { get; } = new();

        public ComandaTransportViewModel()
        {
            _searchService = SearchService.Instance;
            _clientViewmodel = new ClientsViewModel();

            // Data is already loaded by the time this ViewModel is created,
            // so we can immediately populate initial suggestions
            // Load initial suggestions for clients, transportators and locations
            _ = InitializeSuggestionsAsync();
        }

        
        
        private async Task InitializeSuggestionsAsync()
        {
            try
            {
                // Pre-load a few items for each category to show immediate feedback
                var clientTask = LoadAllClientsAsync(new CancellationTokenSource());
                var transportatorTask = LoadAllTransportatorsAsync(new CancellationTokenSource());
                var locationTask = LoadAllLocationsAsync(new CancellationTokenSource(), IncarcareSuggestions);
                var locationTask2 = LoadAllLocationsAsync(new CancellationTokenSource(), DescarcareSuggestions);
                
                await Task.WhenAll(clientTask, transportatorTask, locationTask, locationTask2);
                
                Debug.WriteLine("SearchService: Initial suggestions loaded");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error initializing suggestions: {ex.Message}");
            }
        }

        partial void OnClientChanged(string value)
        {
            // If value is empty, show all clients initially
            if (string.IsNullOrWhiteSpace(value))
            {
                _clientSearchCts?.Cancel();
                _clientSearchCts = new CancellationTokenSource();
                _ = LoadAllClientsAsync(_clientSearchCts);
                return;
            }
            
            // Don't trigger search if it's one of the suggestions (prevents loops)
            if (ClientSuggestions.Contains(value))
            {
                return;
            }

            _clientSearchCts?.Cancel();
            _clientSearchCts = new CancellationTokenSource();
            _ = SearchWithDebounceAsync(value, SearchClientAsync, _clientSearchCts);
        }

        partial void OnTransportatorChanged(string value)
        {
            // If value is empty, show all transportators initially
            if (string.IsNullOrWhiteSpace(value))
            {
                _transportatorSearchCts?.Cancel();
                _transportatorSearchCts = new CancellationTokenSource();
                _ = LoadAllTransportatorsAsync(_transportatorSearchCts);
                return;
            }
            
            // Don't trigger search if it's one of the suggestions (prevents loops)
            if (TransportatorSuggestions.Contains(value))
            {
                return;
            }

            _transportatorSearchCts?.Cancel();
            _transportatorSearchCts = new CancellationTokenSource();
            _ = SearchWithDebounceAsync(value, SearchTransportatorAsync, _transportatorSearchCts);
        }

        partial void OnLocatieIncarcareChanged(string value)
        {
            // If value is empty, show all locations initially
            if (string.IsNullOrWhiteSpace(value))
            {
                _incarcareSearchCts?.Cancel();
                _incarcareSearchCts = new CancellationTokenSource();
                _ = LoadAllLocationsAsync(_incarcareSearchCts, IncarcareSuggestions);
                return;
            }
            
            // Don't trigger search if it's one of the suggestions (prevents loops)
            if (IncarcareSuggestions.Contains(value))
            {
                return;
            }

            _incarcareSearchCts?.Cancel();
            _incarcareSearchCts = new CancellationTokenSource();
            _ = SearchWithDebounceAsync(value, SearchIncarcareAsync, _incarcareSearchCts);
        }

        partial void OnLocatieDescarcareChanged(string value)
        {
            // If value is empty, show all locations initially
            if (string.IsNullOrWhiteSpace(value))
            {
                _descarcareSearchCts?.Cancel();
                _descarcareSearchCts = new CancellationTokenSource();
                _ = LoadAllLocationsAsync(_descarcareSearchCts, DescarcareSuggestions);
                return;
            }
            
            // Don't trigger search if it's one of the suggestions (prevents loops)
            if (DescarcareSuggestions.Contains(value))
            {
                return;
            }

            _descarcareSearchCts?.Cancel();
            _descarcareSearchCts = new CancellationTokenSource();
            _ = SearchWithDebounceAsync(value, SearchDescarcareAsync, _descarcareSearchCts);
        }

        private async Task SearchWithDebounceAsync(string searchText, Func<string, CancellationToken, Task> searchAction, CancellationTokenSource cancellationTokenSource)
        {
            try
            {
                // Minimal delay since we're now searching in memory - much faster!
                await Task.Delay(20, cancellationTokenSource.Token);
                await searchAction(searchText, cancellationTokenSource.Token);
            }
            catch (OperationCanceledException)
            {
                // Expected when search is cancelled
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Search error: {ex.Message}");
            }
        }

        private async Task SearchClientAsync(string searchText, CancellationToken cancellationToken)
        {
            if (string.IsNullOrWhiteSpace(searchText))
            {
                return;
            }

            var results = await _searchService.SearchClientNamesAsync(searchText.Trim());
            
            if (!cancellationToken.IsCancellationRequested)
            {
                ClientSuggestions.Clear();
                foreach (var name in results.Take(10)) // Limit to 10 suggestions
                {
                    ClientSuggestions.Add(name);
                }
            }
        }

        private async Task SearchTransportatorAsync(string searchText, CancellationToken cancellationToken)
        {
            if (string.IsNullOrWhiteSpace(searchText))
            {
                return;
            }

            var results = await _searchService.SearchTransportatorNamesAsync(searchText.Trim());
            
            if (!cancellationToken.IsCancellationRequested)
            {
                TransportatorSuggestions.Clear();
                foreach (var name in results.Take(10)) // Limit to 10 suggestions
                {
                    TransportatorSuggestions.Add(name);
                }
            }
        }

        private async Task SearchIncarcareAsync(string searchText, CancellationToken cancellationToken)
        {
            if (string.IsNullOrWhiteSpace(searchText))
            {
                return;
            }

            var results = await _searchService.SearchLocationAddressesAsync(searchText.Trim());
            
            if (!cancellationToken.IsCancellationRequested)
            {
                IncarcareSuggestions.Clear();
                foreach (var address in results.Take(10)) // Limit to 10 suggestions
                {
                    IncarcareSuggestions.Add(address);
                }
            }
        }

        private async Task SearchDescarcareAsync(string searchText, CancellationToken cancellationToken)
        {
            if (string.IsNullOrWhiteSpace(searchText))
            {
                return;
            }

            var results = await _searchService.SearchLocationAddressesAsync(searchText.Trim());
            
            if (!cancellationToken.IsCancellationRequested)
            {
                DescarcareSuggestions.Clear();
                foreach (var address in results.Take(10)) // Limit to 10 suggestions
                {
                    DescarcareSuggestions.Add(address);
                }
            }
        }

        [RelayCommand]
        private async Task SelectClient(string? clientName)
        {
            if (!string.IsNullOrEmpty(clientName))
            {
                Client = clientName;
                IsClientDropDownOpen = false;
                _clientSearchCts?.Cancel();
                
                // Auto-populate payment terms from client data
                var paymentTerms = await GetTermenPlataAsync(clientName);
                if (!string.IsNullOrEmpty(paymentTerms))
                {
                    TermenPlata = paymentTerms;
                }
            }
        }

        [RelayCommand]
        private void SelectTransportator(string? transportatorName)
        {
            if (!string.IsNullOrEmpty(transportatorName))
            {
                Transportator = transportatorName;
                IsTransportatorDropDownOpen = false;
                _transportatorSearchCts?.Cancel();
            }
        }

        [RelayCommand]
        private void SelectIncarcareLocation(string? location)
        {
            if (!string.IsNullOrEmpty(location))
            {
                LocatieIncarcare = location;
                IsIncarcareDropDownOpen = false;
                _incarcareSearchCts?.Cancel();
            }
        }

        [RelayCommand]
        private void SelectDescarcareLocation(string? location)
        {
            if (!string.IsNullOrEmpty(location))
            {
                LocatieDescarcare = location;
                IsDescarcareDropDownOpen = false;
                _descarcareSearchCts?.Cancel();
            }
        }

        [RelayCommand(CanExecute = nameof(CanSendEmail))]
        private async Task SendEmailAsync()
        {
            if (IsSendingEmail) return;
            
            IsSendingEmail = true;
            
            try
            {
                await GenerateAndSendDocumentAsync();
                await SaveDataInHistoryExcel();
            }
            finally
            {
                IsSendingEmail = false;
            }
        }

        private bool CanSendEmail() {
            return !IsSendingEmail && !string.IsNullOrWhiteSpace(NumarComanda);
        }

        partial void OnIsSendingEmailChanged(bool value)
        {
            SendEmailCommand.NotifyCanExecuteChanged();
        }

        partial void OnNumarComandaChanged(string value)
        {
            SendEmailCommand.NotifyCanExecuteChanged();
        }

        private async Task GenerateAndSendDocumentAsync()
        {
            try
            {
                string projectRoot = AppDomain.CurrentDomain.BaseDirectory;
                string projectDir = FindProjectDirWithDoc(projectRoot);
                string docDir = Path.Combine(projectDir, "doc");
                string mergedTemplatePath = Path.Combine(docDir, "comanda.docx");

                if (!File.Exists(mergedTemplatePath))
                {
                    ShowError($"No template found. Add 'comanda.docx' under: {docDir}", "Template Missing");
                    return;
                }

                string generatedDir = Path.Combine(docDir, "Generated");
                Directory.CreateDirectory(generatedDir);

                string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH.mm.ss");
                var replacements = await BuildCombinedReplacements();
                string mergedOutputPath = Path.Combine(generatedDir, $"CAPAC+Comanda transport - {timestamp}.docx");

                // Generate the document
                GenerateWordDocumentFromTemplate(mergedTemplatePath, mergedOutputPath, replacements);

                // Try to create email
                bool emailCreated = await TryCreateOutlookEmailAsync(mergedOutputPath);

                if (emailCreated)
                {
                    ShowSuccess($"DOCX generated and email draft opened.\n\nDOCX: {mergedOutputPath}", "Ready to Send");
                }
                else
                {
                    // Fallback to opening document directly
                    await TryOpenDocumentAsync(mergedOutputPath);
                }
            }
            catch (Exception ex)
            {
                ShowError($"Failed to generate document.\n\nError: {ex.Message}", "Error");
            }
        }

        private async Task SaveDataInHistoryExcel()
        {
            try
            {
                var excelPath = ClientsViewModel.GetDatabaseExcelPath();
                Directory.CreateDirectory(Path.GetDirectoryName(excelPath)!);

                int numarComandaInt = 0;
                if (!string.IsNullOrWhiteSpace(NumarComanda))
                {
                    var trimmed = NumarComanda.Trim();
                    if (!int.TryParse(trimmed, out numarComandaInt))
                    {
                        Debug.WriteLine($"Warning: could not parse NumarComanda '{trimmed}' to int. Using 0.");
                    }
                }

                

                string clientName = Client?.Trim() ?? string.Empty;
                string camClient = await _searchService.GetClientCameraDeComert(clientName);
                string route = $"{LocatieIncarcare?.Trim() ?? string.Empty} -> {LocatieDescarcare?.Trim() ?? string.Empty}";
                string transportator = Transportator?.Trim() ?? string.Empty;
                DateTime dataTransport = DataIncarcare ?? DateTime.Today;

                var historyEntry = new HistoryTransport(
                    numarComanda: numarComandaInt,
                    clientName: clientName,
                    camClient: camClient,
                    route: route,
                    transportator: transportator,
                    dataTransport: dataTransport
                );
                //need to add laterexcel insert
                await SearchService.Instance.AddHistoryTransportToMemoryAsync(historyEntry);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error saving to history Excel: {ex.Message}");
            }
        }

      

        private static string FindProjectDirWithDoc(string start)
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

        private async Task<bool> TryCreateOutlookEmailAsync(string attachmentPath)
        {
            try
            {
                return await Task.Run(() => CreateOutlookEmailWithAttachment(attachmentPath));
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Email creation failed: {ex}");
                return false;
            }
        }

        private async Task TryOpenDocumentAsync(string documentPath)
        {
            try
            {
                await Task.Run(() => Process.Start(new ProcessStartInfo(documentPath) { UseShellExecute = true }));
                ShowSuccess($"DOCX generated.\n\nDOCX: {documentPath}\n\nOutlook not found or could not open email. The document has been opened directly.", "Success (Manual Email)");
            }
            catch (Exception openEx)
            {
                ShowWarning($"DOCX generated but could not be opened.\n\nDOCX: {documentPath}\nError: {openEx.Message}", "Success (Manual Open Failed)");
            }
        }

        private static void ShowError(string message, string title)
        {
            Application.Current.Dispatcher.Invoke(() =>
                MessageBox.Show(message, title, MessageBoxButton.OK, MessageBoxImage.Error));
        }

        private static void ShowSuccess(string message, string title)
        {
            Application.Current.Dispatcher.Invoke(() =>
                MessageBox.Show(message, title, MessageBoxButton.OK, MessageBoxImage.Information));
        }

        private static void ShowWarning(string message, string title)
        {
            Application.Current.Dispatcher.Invoke(() =>
                MessageBox.Show(message, title, MessageBoxButton.OK, MessageBoxImage.Warning));
        }

        private async Task<Dictionary<string, string>> BuildCombinedReplacements()
        {
            var monedaOptions = new[] { "EUR", "RON", "EUR/MT" };
            var tipOptions = new[] { "TVA", "ALL IN" };
            var tipAdrOptions = new[] { "ADR", "NON-ADR" };
            DateTime today = DateTime.Today;

            var termenPlata = await GetTermenPlataAsync(Transportator);

            return new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "Data azi", today.ToString("dd/MM.yyyy") },
                { "Numar Comanda", NumarComanda?.Trim() ?? string.Empty },
                { "Numar Client", NumarClient?.Trim() ?? string.Empty },
                { "Client", Client?.Trim() ?? string.Empty },
                { "Tarif", Tarif?.Trim() ?? string.Empty },
                { "Moneda", MonedaIndex >= 0 && MonedaIndex < monedaOptions.Length ? monedaOptions[MonedaIndex] : string.Empty },
                { "Tip", TipIndex >= 0 && TipIndex < tipOptions.Length ? tipOptions[TipIndex] : string.Empty },
                { "Transportator", Transportator?.Trim() ?? string.Empty },
                { "Transportator Tarif", TransportatorTarif?.Trim() ?? string.Empty },
                { "Transportator Moneda", TransportatorMonedaIndex >= 0 && TransportatorMonedaIndex < monedaOptions.Length ? monedaOptions[TransportatorMonedaIndex] : string.Empty },
                { "Transportator Tip", TransportatorTipIndex >= 0 && TransportatorTipIndex < tipOptions.Length ? tipOptions[TransportatorTipIndex] : string.Empty },
                { "Data Incarcare", DataIncarcare?.ToString("dd/MM/yyyy") ?? string.Empty },
                { "Data Descarcare", DataDescarcare?.ToString("dd/MM/yyyy") ?? string.Empty },
                { "Produs", Produs?.Trim() ?? string.Empty },
                { "Cantitate", Cantitate?.Trim() ?? string.Empty },
                { "Tip ADR", TipAdrIndex >= 0 && TipAdrIndex < tipAdrOptions.Length ? tipAdrOptions[TipAdrIndex] : string.Empty },
                { "Clasa", Clasa?.Trim() ?? string.Empty },
                { "UN", Un?.Trim() ?? string.Empty },
                { "Numar Inmatriculare", NumarInmatriculare?.Trim() ?? string.Empty },
                { "Adresa Incarcare", LocatieIncarcare?.Trim() ?? string.Empty },
                { "Adresa Descarcare", LocatieDescarcare?.Trim() ?? string.Empty },
                { "Termen Plata", TermenPlata?.Trim() ?? string.Empty }
            };
        }

        private async Task<string> GetTermenPlataAsync(string transportatorName)
        {
            try
            {
                var tranportatorFound = await SearchService.Instance.GetTransportatorByNameAsync(transportatorName);
                return tranportatorFound?.TermenulDePlata ?? string.Empty;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error getting termen plata for client {transportatorName}: {ex.Message}");
                return string.Empty;
            }
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

        private async Task LoadAllClientsAsync(CancellationTokenSource cancellationTokenSource)
        {
            try
            {
                var results = await _searchService.GetAllClientNamesAsync();
                
                if (!cancellationTokenSource.Token.IsCancellationRequested)
                {
                    ClientSuggestions.Clear();
                    foreach (var name in results.Take(15)) // Show more items when showing all
                    {
                        ClientSuggestions.Add(name);
                    }
                }
            }
            catch (OperationCanceledException)
            {
                // Expected when cancelled
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Load all clients error: {ex.Message}");
            }
        }

        private async Task LoadAllTransportatorsAsync(CancellationTokenSource cancellationTokenSource)
        {
            try
            {
                var results = await _searchService.GetAllTransportatorNamesAsync();
                
                if (!cancellationTokenSource.Token.IsCancellationRequested)
                {
                    TransportatorSuggestions.Clear();
                    foreach (var name in results.Take(15)) // Show more items when showing all
                    {
                        TransportatorSuggestions.Add(name);
                    }
                }
            }
            catch (OperationCanceledException)
            {
                // Expected when cancelled
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Load all transportators error: {ex.Message}");
            }
        }

        private async Task LoadAllLocationsAsync(CancellationTokenSource cancellationTokenSource, ObservableCollection<string> targetCollection)
        {
            try
            {
                var results = await _searchService.GetAllLocationAddressesAsync();
                
                if (!cancellationTokenSource.Token.IsCancellationRequested)
                {
                    targetCollection.Clear();
                    foreach (var address in results.Take(15)) // Show more items when showing all
                    {
                        targetCollection.Add(address);
                    }
                }
            }
            catch (OperationCanceledException)
            {
                // Expected when cancelled
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Load all locations error: {ex.Message}");
            }
        }
    }
}