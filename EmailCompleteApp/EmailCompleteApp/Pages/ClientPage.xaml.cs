using System;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using EmailCompleteApp.Models;
using EmailCompleteApp.ViewModels;

namespace EmailCompleteApp.Pages
{
    public partial class ClientPage : UserControl
    {
        private ClientsViewModel? ViewModel => DataContext as ClientsViewModel;

        public ClientPage()
        {
            InitializeComponent();
            DataContext = new ClientsViewModel();
        }

        private void EditClient_Click(object sender, RoutedEventArgs e)
        {
            if (ClientDataGrid.SelectedItem is Client client)
            {
                ViewModel?.EditClient(client);
            }
        }

        private void DeleteClient_Click(object sender, RoutedEventArgs e)
        {
            if (ClientDataGrid.SelectedItem is Client client)
            {
                _ = ViewModel?.DeleteClientCommand.ExecuteAsync(client);
            }
        }

        private void EditTransportator_Click(object sender, RoutedEventArgs e)
        {
            if (TransportatorDataGrid.SelectedItem is Transportator transportator)
            {
                ViewModel?.EditTransportator(transportator);
            }
        }

        private void DeleteTransportator_Click(object sender, RoutedEventArgs e)
        {
            if (TransportatorDataGrid.SelectedItem is Transportator transportator)
            {
                _ = ViewModel?.DeleteTransportatorCommand.ExecuteAsync(transportator);
            }
        }
    }

    /// <summary>
    /// Converter that inverts boolean values for visibility binding
    /// </summary>
    public class InvertedBooleanToVisibilityConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is bool boolValue)
            {
                return boolValue ? Visibility.Collapsed : Visibility.Visible;
            }
            return Visibility.Visible;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is Visibility visibility)
            {
                return visibility != Visibility.Visible;
            }
            return false;
        }
    }
}
