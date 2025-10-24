using System;
using System.IO;
using System.Linq;
using System.Windows;
using ClosedXML.Excel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using EmailCompleteApp.Models;
using EmailCompleteApp.Services;

namespace EmailCompleteApp.ViewModels;

public partial class ClientsViewModel : ObservableObject
{

    SearchService _searchService = SearchService.Instance;
    [ObservableProperty] private int _selectedTabIndex = 0;

    [ObservableProperty] private string _clientName = string.Empty;
    [ObservableProperty] private string _clientAddress = string.Empty;
    [ObservableProperty] private string _clientBank = string.Empty;
    [ObservableProperty] private string _clientIban = string.Empty;
    [ObservableProperty] private string _clientVatNumber = string.Empty;
    [ObservableProperty] private string _clientCameraDeComert = string.Empty;
    [ObservableProperty] private string _clientTermenPlata = string.Empty;

    [ObservableProperty] private string _transportatorName = string.Empty;
    [ObservableProperty] private string _transportatorAdresa = string.Empty;
    [ObservableProperty] private string _transportatorContBancar = string.Empty;
    [ObservableProperty] private string _transportatorIban = string.Empty;
    [ObservableProperty] private string _transportatorVatNumber = string.Empty;
    [ObservableProperty] private string _transportatorCameraDeComert = string.Empty;
    [ObservableProperty] private string _transportatorTermenPlata = string.Empty;

    [RelayCommand]
    private async void Insert()
    {
        try
        {
            var excelPath = GetDatabaseExcelPath();
            Directory.CreateDirectory(Path.GetDirectoryName(excelPath)!);

            if (SelectedTabIndex == 0)
            {
                if (string.IsNullOrWhiteSpace(ClientName)) { ShowWarn("Name is required."); return; }
                if (string.IsNullOrWhiteSpace(ClientAddress)) { ShowWarn("Address is required."); return; }

                // Get next client ID
                var allClients = await _searchService.GetAllClientsAsync();
                int nextId = allClients.Count > 0 ? allClients.Max(c => c.Id) + 1 : 1;

                var client = new Client(nextId, ClientName, ClientAddress, ClientBank, ClientIban, ClientVatNumber, ClientCameraDeComert, ClientTermenPlata);


                AppendClientToExcel(excelPath, "Clients", client);
                
                // Add to in-memory list (no need to reload from Excel)
                await _searchService.AddClientToMemoryAsync(client);
                
                MessageBox.Show($"Client inserted to Excel:\n{excelPath}", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
                
                ClientName = ClientAddress = ClientBank = ClientIban = ClientVatNumber = ClientCameraDeComert = ClientTermenPlata = string.Empty;
            }
            else
            {
                if (string.IsNullOrWhiteSpace(TransportatorName)) { ShowWarn("Name is required."); return; }
                if (string.IsNullOrWhiteSpace(TransportatorAdresa)) { ShowWarn("Address is required."); return; }
                
                // Get next transportator ID
                var allTransportators = await _searchService.GetAllTransportatorsAsync();
                int nextId = allTransportators.Count > 0 ? allTransportators.Max(t => t.Id) + 1 : 1;
                
                var transportator = new Transportator(nextId, TransportatorName, TransportatorAdresa, TransportatorContBancar, TransportatorIban, TransportatorVatNumber, TransportatorCameraDeComert, TransportatorTermenPlata);
                transportator.Id = nextId;
                
                AppendTransportatorToExcel(excelPath, "Transportators", transportator);
                
                // Add to in-memory list (no need to reload from Excel)
                await _searchService.AddTransportatorToMemoryAsync(transportator);
                
                MessageBox.Show($"Transportator inserted to Excel:\n{excelPath}", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
                
                TransportatorName = TransportatorAdresa = TransportatorContBancar = TransportatorIban = TransportatorVatNumber = TransportatorCameraDeComert = TransportatorTermenPlata = string.Empty;
            }
        }
        catch (ArgumentException ex)
        {
            ShowWarn(ex.Message);
        }
        catch (IOException ioEx)
        {
            MessageBox.Show($"Could not write the Excel file. It may be open in another program.\nDetails: {ioEx.Message}", "File In Use", MessageBoxButton.OK, MessageBoxImage.Warning);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private static void ShowWarn(string msg) => MessageBox.Show(msg, "Validation Error", MessageBoxButton.OK, MessageBoxImage.Warning);

    public static string GetDatabaseExcelPath()
    {
        string baseDir = AppDomain.CurrentDomain.BaseDirectory;
        string? current = baseDir;
        for (int i = 0; i < 6 && current != null; i++)
        {
            string docDir = Path.Combine(current, "doc");
            if (Directory.Exists(docDir))
                return Path.Combine(docDir, "database.xlsx");
            current = Directory.GetParent(current)?.FullName;
        }
        var docs = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        return Path.Combine(docs, "AutoEmailCompletion", "database.xlsx");
    }

    private static void EnsureHeader(IXLWorksheet ws)
    {
        ws.Cell(1, 1).Value = "ID";
        ws.Cell(1, 2).Value = "Name";
        ws.Cell(1, 3).Value = "Address";
        ws.Cell(1, 4).Value = "Bank";
        ws.Cell(1, 5).Value = "IBAN";
        ws.Cell(1, 6).Value = "VAT";
        ws.Cell(1, 7).Value = "Camera de Comert";
        ws.Cell(1, 8).Value = "Termenul de plata";
        ws.Cell(1, 9).Value = "CreatedAt";
        ws.Row(1).Style.Font.Bold = true;
    }

    private static void AppendClientToExcel(string filePath, string sheetName, Client client)
    {
        if (!File.Exists(filePath) || new FileInfo(filePath).Length == 0)
        {
            using var initWb = new XLWorkbook();
            var initWs = initWb.Worksheets.Add(sheetName);
            EnsureHeader(initWs);
            initWs.Columns().AdjustToContents();
            initWb.SaveAs(filePath);
        }

        XLWorkbook? wb = null;
        try
        {
            wb = new XLWorkbook(filePath);
        }
        catch
        {
            using var recreate = new XLWorkbook();
            var wsNew = recreate.Worksheets.Add(sheetName);
            EnsureHeader(wsNew);
            recreate.SaveAs(filePath);
            wb = new XLWorkbook(filePath);
        }

        using (wb)
        {
            var wsExisting = wb.Worksheets.FirstOrDefault(s => s.Name.Equals(sheetName, StringComparison.OrdinalIgnoreCase))
                           ?? wb.Worksheets.Add(sheetName);
            if (wsExisting.Cell(1, 1).GetString().Length == 0) EnsureHeader(wsExisting);
            var lastRow = wsExisting.LastRowUsed()?.RowNumber() ?? 1;
            var targetRow = lastRow >= 1 ? lastRow + 1 : 2;
            wsExisting.Cell(targetRow, 1).Value = client.Id;
            wsExisting.Cell(targetRow, 2).Value = client.Name;
            wsExisting.Cell(targetRow, 3).Value = client.Address;
            wsExisting.Cell(targetRow, 4).Value = client.Bank;
            wsExisting.Cell(targetRow, 5).Value = client.IBAN;
            wsExisting.Cell(targetRow, 6).Value = client.VATNumber;
            wsExisting.Cell(targetRow, 7).Value = client.CameraDeComert;
            wsExisting.Cell(targetRow, 8).Value = client.TermenulDePlata;
            wsExisting.Cell(targetRow, 9).Value = DateTime.Now;
            wsExisting.Columns().AdjustToContents();
            wb.Save();
        }
    }

    private static void AppendTransportatorToExcel(string filePath, string sheetName, Transportator transportator)
    {
        if (!File.Exists(filePath) || new FileInfo(filePath).Length == 0)
        {
            using var initWb = new XLWorkbook();
            var initWs = initWb.Worksheets.Add(sheetName);
            EnsureHeader(initWs);
            initWs.Columns().AdjustToContents();
            initWb.SaveAs(filePath);
        }

        XLWorkbook? wb = null;
        try
        {
            wb = new XLWorkbook(filePath);
        }
        catch
        {
            using var recreate = new XLWorkbook();
            var wsNew = recreate.Worksheets.Add(sheetName);
            EnsureHeader(wsNew);
            recreate.SaveAs(filePath);
            wb = new XLWorkbook(filePath);
        }

        using (wb)
        {
            var wsExisting = wb.Worksheets.FirstOrDefault(s => s.Name.Equals(sheetName, StringComparison.OrdinalIgnoreCase))
                           ?? wb.Worksheets.Add(sheetName);
            if (wsExisting.Cell(1, 1).GetString().Length == 0) EnsureHeader(wsExisting);
            var lastRow = wsExisting.LastRowUsed()?.RowNumber() ?? 1;
            var targetRow = lastRow >= 1 ? lastRow + 1 : 2;
            wsExisting.Cell(targetRow, 1).Value = transportator.Id;
            wsExisting.Cell(targetRow, 2).Value = transportator.Name;
            wsExisting.Cell(targetRow, 3).Value = transportator.Address;
            wsExisting.Cell(targetRow, 4).Value = transportator.Bank;
            wsExisting.Cell(targetRow, 5).Value = transportator.IBAN;
            wsExisting.Cell(targetRow, 6).Value = transportator.VATNumber;
            wsExisting.Cell(targetRow, 7).Value = transportator.CameraDeComert;
            wsExisting.Cell(targetRow, 8).Value = transportator.TermenulDePlata;
            wsExisting.Cell(targetRow, 9).Value = DateTime.Now;
            wsExisting.Columns().AdjustToContents();
            wb.Save();
        }
    }
}
