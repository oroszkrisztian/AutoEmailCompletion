using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using ClosedXML.Excel;
using EmailCompleteApp.Models;

namespace EmailCompleteApp.Services
{
    public class SearchService
    {
        public async Task<List<string>> SearchClientNamesAsync(string searchText)
        {
            return await SearchInSheetAsync(searchText, "Clients", 1); // Column 1 for names
        }

        public async Task<List<string>> SearchTransportatorNamesAsync(string searchText)
        {
            // Only search Transportators sheet, no fallback to Clients
            return await SearchInSheetAsync(searchText, "Transportators", 1);
        }

        public async Task<List<string>> SearchLocationAddressesAsync(string searchText)
        {
            return await SearchInSheetAsync(searchText, "Locations", 2); // Column 2 for addresses in Locations sheet
        }

        private async Task<List<string>> SearchInSheetAsync(string searchText, string sheetName, int columnIndex)
        {
            if (string.IsNullOrWhiteSpace(searchText) || searchText.Length < 1)
                return new List<string>();

            return await Task.Run(() =>
            {
                var results = new List<string>();
                string excelPath = GetDatabaseExcelPath();
                
                if (!File.Exists(excelPath))
                    return results;

                try
                {
                    using var workbook = new XLWorkbook(excelPath);
                    var targetSheet = workbook.Worksheets.FirstOrDefault(s => 
                        s.Name.Equals(sheetName, StringComparison.OrdinalIgnoreCase));
                    
                    if (targetSheet == null)
                        return results;

                    var lastRow = targetSheet.LastRowUsed()?.RowNumber() ?? 1;
                    
                    for (int row = 2; row <= lastRow && results.Count < 10; row++)
                    {
                        try
                        {
                            var cellValue = targetSheet.Cell(row, columnIndex).GetString();
                            
                            if (!string.IsNullOrWhiteSpace(cellValue) && 
                                cellValue.Contains(searchText, StringComparison.OrdinalIgnoreCase))
                            {
                                if (!results.Contains(cellValue))
                                {
                                    results.Add(cellValue);
                                }
                            }
                        }
                        catch
                        {
                            continue;
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"An error occurred while reading the Excel file: {ex.Message}", 
                                   "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
                
                return results;
            });
        }

        private static string GetDatabaseExcelPath()
        {
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;

            string? current = baseDir;
            for (int i = 0; i < 6 && current != null; i++)
            {
                string docDir = Path.Combine(current, "doc");
                if (Directory.Exists(docDir))
                {
                    return Path.Combine(docDir, "database.xlsx");
                }
                current = Directory.GetParent(current)?.FullName;
            }

            var docs = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            return Path.Combine(docs, "AutoEmailCompletion", "database.xlsx");
        }

        public async Task<List<string>> GetAllClientNamesAsync()
        {
            return await Task.Run(() =>
            {
                var results = new List<string>();
                string excelPath = GetDatabaseExcelPath();
                
                if (!File.Exists(excelPath))
                    return results;

                try
                {
                    using var workbook = new XLWorkbook(excelPath);
                    var clientSheet = workbook.Worksheets.FirstOrDefault(s => 
                        s.Name.Equals("Clients", StringComparison.OrdinalIgnoreCase));
                    
                    if (clientSheet == null)
                        return results;

                    var lastRow = clientSheet.LastRowUsed()?.RowNumber() ?? 1;
                    
                    for (int row = 2; row <= lastRow; row++)
                    {
                        try
                        {
                            var clientName = clientSheet.Cell(row, 1).GetString();
                            
                            if (!string.IsNullOrWhiteSpace(clientName))
                            {
                                if (!results.Contains(clientName))
                                {
                                    results.Add(clientName);
                                }
                            }
                        }
                        catch
                        {
                            continue;
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"An error occurred while reading the Excel file: {ex.Message}", 
                                   "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
                
                return results;
            });
        }
    }
}