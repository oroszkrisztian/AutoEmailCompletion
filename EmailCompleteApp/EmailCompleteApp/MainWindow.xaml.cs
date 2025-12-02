using System.Windows;
using System.Windows.Input;
using System.Windows.Media;
using EmailCompleteApp.Pages;

namespace EmailCompleteApp
{
    public partial class MainWindow : Window
    {
        private enum SelectedButton { None, Client, Location, Comanda, Istoric }
        private SelectedButton _selected = SelectedButton.None;

        public MainWindow()
        {
            InitializeComponent();
        }

        private void ClientButton_Click(object sender, RoutedEventArgs e)
        {
            // Switch to Client page
            MainContentArea.Content = new ClientPage();
            _selected = SelectedButton.Client;
            UpdateButtonVisuals();
        }

        private void LocationButton_Click(object sender, RoutedEventArgs e)
        {
            MainContentArea.Content = new LocationPage();
            _selected = SelectedButton.Location;
            UpdateButtonVisuals();
        }

        private void ComandaTransport_Click(object sender, RoutedEventArgs e)
        {
            MainContentArea.Content = new ComandaTransport();
            _selected = SelectedButton.Comanda;
            UpdateButtonVisuals();
        }

        private void ComandaTransportIstorict_Click(object sender, RoutedEventArgs e)
        {
            MainContentArea.Content = new HistoryPage();
            _selected = SelectedButton.Istoric;
            UpdateButtonVisuals();
        }

        // Mouse Enter Event Handlers
        private void ClientButton_MouseEnter(object sender, MouseEventArgs e)
        {
            if (ClientButtonPage.Background != new SolidColorBrush(Color.FromRgb(100, 150, 255)))
            {
                ClientButtonPage.Background = new SolidColorBrush(Color.FromRgb(100, 150, 255));
                ClientButtonPage.Foreground = Brushes.White;
            }
        }

        private void LocationButton_MouseEnter(object sender, MouseEventArgs e)
        {
            if (LocationButtonPage.Background != new SolidColorBrush(Color.FromRgb(100, 150, 255)))
            {
                LocationButtonPage.Background = new SolidColorBrush(Color.FromRgb(100, 150, 255));
                LocationButtonPage.Foreground = Brushes.White;
            }
        }

        private void ComandaTransportButton_MouseEnter(object sender, MouseEventArgs e)
        {
            if (ComandaTransportButton.Background != new SolidColorBrush(Color.FromRgb(100, 150, 255)))
            {
                ComandaTransportButton.Background = new SolidColorBrush(Color.FromRgb(100, 150, 255));
                ComandaTransportButton.Foreground = Brushes.White;
            }
        }

        private void ComandaTransportIstorictButton_MouseEnter(object sender, MouseEventArgs e)
        {
            if (ComandaTransportIstoricButton.Background != new SolidColorBrush(Color.FromRgb(100, 150, 255)))
            {
                ComandaTransportIstoricButton.Background = new SolidColorBrush(Color.FromRgb(100, 150, 255));
                ComandaTransportIstoricButton.Foreground = Brushes.White;
            }
        }

        // Mouse Leave Event Handlers
        private void ClientButton_MouseLeave(object sender, MouseEventArgs e)
        {
            if (MainContentArea.Content is ClientPage)
            {
                // Keep selected state - do nothing
                return;
            }
            else
            {
                // Reset to default state
                ClientButtonPage.Background = Brushes.White;
                ClientButtonPage.Foreground = new SolidColorBrush(Color.FromRgb(39, 37, 55));
            }
        }

        private void LocationButton_MouseLeave(object sender, MouseEventArgs e)
        {
            // If the currently displayed content is the Location page, keep the selected state
            if (MainContentArea.Content is LocationPage)
            {
                return;
            }

            // Otherwise reset to default
            LocationButtonPage.Background = Brushes.White;
            LocationButtonPage.Foreground = new SolidColorBrush(Color.FromRgb(39, 37, 55));
        }

        private void ComandaTransport_MouseLeave(object sender, MouseEventArgs e)
        {
            // Only reset if this button is not the currently selected one
            if (MainContentArea.Content is ComandaTransport)
            {
                // Keep selected state - do nothing
                return;
            }
            else
            {
                // Reset to default state
                ComandaTransportButton.Background = Brushes.White;
                ComandaTransportButton.Foreground = new SolidColorBrush(Color.FromRgb(39, 37, 55));
            }
        }

        private void ComandaTransportIstoric_MouseLeave(object sender, MouseEventArgs e)
        {
            if (MainContentArea.Content is HistoryPage)
            {
                return;
            }
            else
            {
                ComandaTransportIstoricButton.Background = Brushes.White;
                ComandaTransportIstoricButton.Foreground = new SolidColorBrush(Color.FromRgb(39, 37, 55));
            }
        }

        private void UpdateButtonVisuals()
        {
            // Reset all to default first
            ClientButtonPage.Background = Brushes.White;
            ClientButtonPage.Foreground = new SolidColorBrush(Color.FromRgb(39, 37, 55));

            ComandaTransportButton.Background = Brushes.White;
            ComandaTransportButton.Foreground = new SolidColorBrush(Color.FromRgb(39, 37, 55));

            LocationButtonPage.Background = Brushes.White;
            LocationButtonPage.Foreground = new SolidColorBrush(Color.FromRgb(39, 37, 55));

            ComandaTransportIstoricButton.Background = Brushes.White;
            ComandaTransportIstoricButton.Foreground = new SolidColorBrush(Color.FromRgb(39, 37, 55));

            // Apply selected visuals
            var selectedBrush = new SolidColorBrush(Color.FromRgb(100, 150, 255));
            switch (_selected)
            {
                case SelectedButton.Client:
                    ClientButtonPage.Background = selectedBrush;
                    ClientButtonPage.Foreground = Brushes.White;
                    break;
                case SelectedButton.Location:
                    LocationButtonPage.Background = selectedBrush;
                    LocationButtonPage.Foreground = Brushes.White;
                    break;
                case SelectedButton.Comanda:
                    ComandaTransportButton.Background = selectedBrush;
                    ComandaTransportButton.Foreground = Brushes.White;
                    break;
                case SelectedButton.Istoric:
                    ComandaTransportIstoricButton.Background = selectedBrush;
                    ComandaTransportIstoricButton.Foreground = Brushes.White;
                    break;
            }
        }
    }
}