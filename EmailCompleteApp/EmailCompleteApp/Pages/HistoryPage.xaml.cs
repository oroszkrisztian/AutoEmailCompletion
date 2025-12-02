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

        }

        private void HistoryListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (HistoryListBox.SelectedItem is HistoryTransport selectedItem)
            {
                _viewModel.OpenDocument(selectedItem.NumarComanda.ToString());
                MessageBox.Show($"Clicked on history item: {selectedItem.NumarComanda}");
            }
            HistoryListBox.UnselectAll();
        }
    }
}