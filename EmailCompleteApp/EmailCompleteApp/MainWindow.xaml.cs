using System.Windows;
using System.Windows.Input;
using System.Windows.Media;
using EmailCompleteApp.Pages;
using EmailCompleteApp.ViewModels;

namespace EmailCompleteApp
{

    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            DataContext = new MainViewModel();
        }

        private void TitleBar_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left)
                this.DragMove();
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void ClientButton_Click(object sender, RoutedEventArgs e)
        {
            // Switch to Cleint page
            MainContentArea.Content = new ClientPage();

            // Update button styles to show active state
            ResetButtonStyles();
            ClientButtonPage.Background = new SolidColorBrush(Color.FromRgb(100, 150, 255));
            ClientButtonPage.Foreground = Brushes.White;
        }

        private void LocationButton_Click(object sender, RoutedEventArgs e)
        {
            MainContentArea.Content = new LocationPage();
            ResetButtonStyles();
            LocationButtonPage.Background = new SolidColorBrush(Color.FromRgb(100, 150, 255));
            LocationButtonPage.Foreground= Brushes.White;
        }

        private void ComandaTransport_Click(object sender, RoutedEventArgs e)
        {
            MainContentArea.Content = new ComandaTransport();
            
            // Update button styles to show active state
            ResetButtonStyles();
            ComandaTransportButton.Background = new SolidColorBrush(Color.FromRgb(100, 150, 255));
            ComandaTransportButton.Foreground = Brushes.White;
        }

       
        private void ResetButtonStyles()
        {
            // Reset all buttons to default style
            ClientButtonPage.Background = Brushes.White;
            ClientButtonPage.Foreground = new SolidColorBrush(Color.FromRgb(39, 37, 55));

            ComandaTransportButton.Background = Brushes.White;
            ComandaTransportButton.Foreground = new SolidColorBrush(Color.FromRgb(39, 37, 55));
            
           
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

            if (MainContentArea.Content is ClientPage)
            {
                // Keep selected state - do nothing
                return;
            }
            else
            {
                // Reset to default state
                LocationButtonPage.Background = Brushes.White;
                LocationButtonPage.Foreground = new SolidColorBrush(Color.FromRgb(39, 37, 55));
            }
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

        
        // Minimize Button Event Handlers
        private void MinimizeButton_Click(object sender, RoutedEventArgs e)
        {
            this.WindowState = WindowState.Minimized;
        }

        private void MinimizeButton_MouseEnter(object sender, MouseEventArgs e)
        {
            MinimizeButton.Background = new SolidColorBrush(Color.FromRgb(100, 100, 100)); // Gray hover effect
        }

        private void MinimizeButton_MouseLeave(object sender, MouseEventArgs e)
        {
            MinimizeButton.Background = Brushes.Transparent;
        }

        // Close Button Mouse Event Handlers
        private void CloseButton_MouseEnter(object sender, MouseEventArgs e)
        {
            CloseButton.Background = new SolidColorBrush(Color.FromRgb(255, 68, 68)); // #FF4444
        }

        private void CloseButton_MouseLeave(object sender, MouseEventArgs e)
        {
            CloseButton.Background = Brushes.Transparent;
        }
    }
}