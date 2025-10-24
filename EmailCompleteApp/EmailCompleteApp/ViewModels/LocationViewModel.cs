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

public partial class LocationViewModel : ObservableObject
{
    private readonly SearchService _searchService = SearchService.Instance;
    
    [ObservableProperty] private string _locationName = string.Empty;
    [ObservableProperty] private string _locationAddress = string.Empty;
    [ObservableProperty] private string _locationCity = string.Empty;

    [RelayCommand]
    private async void Save()
    {
        try
        {
            if (string.IsNullOrWhiteSpace(LocationName)) { Warn("Name is required."); return; }
            if (string.IsNullOrWhiteSpace(LocationAddress)) { Warn("Address is required."); return; }
            if (string.IsNullOrWhiteSpace(LocationCity)) { Warn("City is required."); return; }

            // Get next location ID
            var allLocations = await _searchService.GetAllLocationsAsync();
            int nextId = allLocations.Count > 0 ? allLocations.Max(l => l.Id) + 1 : 1;
            var location = new Location(nextId, LocationName, LocationAddress, LocationCity);
            var excelPath = GetLocationsExcelPath();
            Directory.CreateDirectory(Path.GetDirectoryName(excelPath)!);
            AppendLocationsToExcel(excelPath, location);
            
            // Add to in-memory list (no need to reload from Excel)
            await _searchService.AddLocationToMemoryAsync(location);
            
            MessageBox.Show($"Location inserted to Excel:\n{excelPath}", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
            
            LocationName = string.Empty;
            LocationAddress = string.Empty;
            LocationCity = string.Empty;
        }
        catch (ArgumentException ex)
        {
            Warn(ex.Message);
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

    private static void Warn(string msg) => MessageBox.Show(msg, "Validation Error", MessageBoxButton.OK, MessageBoxImage.Warning);

    private static string GetLocationsExcelPath()
    {
        string baseDir = AppDomain.CurrentDomain.BaseDirectory;
        string? current = baseDir;
        for (int i = 0; i < 6 && current != null; i++)
        {
            string docDir = Path.Combine(current, "doc");
            if (Directory.Exists(docDir))
            {
                string preferred = Path.Combine(docDir, "database.xlsx");
                string typo = Path.Combine(docDir, "database.xlxs");
                if (File.Exists(typo) && !File.Exists(preferred)) return typo;
                return preferred;
            }
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
        ws.Cell(1, 4).Value = "City";
        ws.Cell(1, 5).Value = "CreatedAt";
        ws.Row(1).Style.Font.Bold = true;
    }

    private static void AppendLocationsToExcel(string filePath, Location location)
    {
        const string sheetName = "Locations";
        if (!File.Exists(filePath) || new FileInfo(filePath).Length == 0)
        {
            using var initWb = new XLWorkbook();
            var initWs = initWb.Worksheets.Add(sheetName);
            EnsureHeader(initWs);
            initWs.Columns().AdjustToContents();
            initWb.SaveAs(filePath);
        }
        XLWorkbook? wb = null;
        try { wb = new XLWorkbook(filePath); }
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
            var wsExisting = wb.Worksheets.FirstOrDefault(s => s.Name.Equals(sheetName, StringComparison.OrdinalIgnoreCase)) ?? wb.Worksheets.Add(sheetName);
            if (wsExisting.Cell(1, 1).GetString().Length == 0) EnsureHeader(wsExisting);
            var lastRow = wsExisting.LastRowUsed()?.RowNumber() ?? 1;
            var targetRow = lastRow >= 1 ? lastRow + 1 : 2;
            wsExisting.Cell(targetRow, 1).Value = location.Id;
            wsExisting.Cell(targetRow, 2).Value = location.Name;
            wsExisting.Cell(targetRow, 3).Value = location.Address;
            wsExisting.Cell(targetRow, 4).Value = location.City;
            wsExisting.Cell(targetRow, 5).Value = DateTime.Now;
            wsExisting.Columns().AdjustToContents();
            wb.Save();
        }
    }
}
