using ClosedXML;
using EmailCompleteApp.Models;
using EmailCompleteApp.ViewModels;
using System;
using System.Collections.Generic;
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
    /// <summary>
    /// Interaction logic for HistoryPage.xaml
    /// </summary>
    public partial class HistoryPage : UserControl
    {
        private HistoryPageViewModel? _viewModel;

        public HistoryPage()
        {
            InitializeComponent();
            DataContext = new HistoryPageViewModel();
            _viewModel = DataContext as HistoryPageViewModel;

            if (_viewModel != null)
            {
                _viewModel.EditRequested += OnEditRequested;
            }
        }

        private void OnEditRequested(HistoryTransport historyItem)
        {
            // Navigate to ComandaTransport page with edit data
            var mainWindow = Window.GetWindow(this) as MainWindow;
            if (mainWindow != null)
            {
                var comandaPage = new ComandaTransport(historyItem);
                mainWindow.NavigateToComandaTransport(comandaPage);
            }
        }

        private void HistoryDataGrid_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (_viewModel != null && HistoryDataGrid.SelectedItem is HistoryTransport selectedHistory)
            {
                if (!string.IsNullOrWhiteSpace(selectedHistory.NumarComanda))
                {
                    _viewModel.OpenDocument(selectedHistory.NumarComanda);
                }
            }
        }

        private void OpenDocument_Click(object sender, RoutedEventArgs e)
        {
            if (HistoryDataGrid.SelectedItem is HistoryTransport historyItem)
            {
                if (!string.IsNullOrWhiteSpace(historyItem.NumarComanda))
                {
                    _viewModel?.OpenDocument(historyItem.NumarComanda);
                }
            }
        }

        private void EditOrder_Click(object sender, RoutedEventArgs e)
        {
            if (HistoryDataGrid.SelectedItem is HistoryTransport historyItem)
            {
                _viewModel?.EditOrder(historyItem);
            }
        }
    }
}