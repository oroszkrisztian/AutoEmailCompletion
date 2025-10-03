using ClosedXML.Excel;
using EmailCompleteApp.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace EmailCompleteApp.Pages
{
    
    public partial class LocationPage : UserControl
    {
        public LocationPage()
        {
            InitializeComponent();
        }

        private void OnSaveLocationClick(object sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(LocationName.Text))
                {
                    MessageBox.Show("Name is required.", "Validation Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                    LocationName.Focus();
                    return;
                }

                if (string.IsNullOrWhiteSpace(LocationAddress.Text))
                {
                    MessageBox.Show("Address is required.", "Validation Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                    LocationAddress.Focus();
                    return;
                }

                Location location = new Location(
                    LocationName.Text,
                    LocationAddress.Text
                    
                );

                

                var excelPath = GetLocationsExcelPath();
                System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(excelPath)!);




                AppendLocationsToExcel(excelPath, location);

                MessageBox.Show($"Location inserted to Excel:\n{excelPath}", "Success", MessageBoxButton.OK, MessageBoxImage.Information);

                LocationName.Text = string.Empty;
                LocationAddress.Text = string.Empty;
            }
            catch (ArgumentException ex)
            {
                MessageBox.Show(ex.Message, "Validation Error", MessageBoxButton.OK, MessageBoxImage.Warning);
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

        private string GetLocationsExcelPath() 
        {
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;

            string? current = baseDir;
            for (int i = 0; i < 6 && current != null; i++)
            {
                string docDir = System.IO.Path.Combine(current, "doc");
                if (Directory.Exists(docDir))
                {
                    string preferred = System.IO.Path.Combine(docDir, "database.xlsx");
                    string typo = System.IO.Path.Combine(docDir, "database.xlxs");
                    if (File.Exists(typo) && !File.Exists(preferred))
                    {
                        return typo;
                    }
                    return preferred;
                }
                current = Directory.GetParent(current)?.FullName;
            }

            var docs = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            return System.IO.Path.Combine(docs, "AutoEmailCompletion", "database.xlsx");
        }

        private static void EnsureHeader(IXLWorksheet ws)
        {
            ws.Cell(1, 1).Value = "Name";
            ws.Cell(1, 2).Value = "Address";
            ws.Cell(1, 8).Value = "CreatedAt";
            ws.Row(1).Style.Font.Bold = true;
        }

        private void AppendLocationsToExcel(string filePath, Location location)
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

                if (wsExisting.Cell(1, 1).GetString().Length == 0)
                {
                    EnsureHeader(wsExisting);
                }

                var lastRow = wsExisting.LastRowUsed()?.RowNumber() ?? 1;
                var targetRow = lastRow >= 1 ? lastRow + 1 : 2;

                wsExisting.Cell(targetRow, 1).Value = location.firmName;
                wsExisting.Cell(targetRow, 2).Value = location.address;
                wsExisting.Cell(targetRow, 8).Value = DateTime.Now;
                wsExisting.Columns().AdjustToContents();

                wb.Save();
            }
        }
    }
}
