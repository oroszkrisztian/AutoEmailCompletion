using System.Windows;
using System.Windows.Input;
using System.Windows.Media;
using EmailCompleteApp.Pages;

namespace EmailCompleteApp
{

    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
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

        private void EmailType2Button_Click(object sender, RoutedEventArgs e)
        {
            // Switch to Email Type 1 page
            MainContentArea.Content = new EmailType2Page();
            
            // Update button styles to show active state
            ResetButtonStyles();
            EmailType2Button.Background = new SolidColorBrush(Color.FromRgb(100, 150, 255));
            EmailType2Button.Foreground = Brushes.White;
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
            EmailType2Button.Background = Brushes.White;
            EmailType2Button.Foreground = new SolidColorBrush(Color.FromRgb(39, 37, 55));

            ComandaTransportButton.Background = Brushes.White;
            ComandaTransportButton.Foreground = new SolidColorBrush(Color.FromRgb(39, 37, 55));
            
           
        }

        // Mouse Enter Event Handlers
        private void EmailType2Button_MouseEnter(object sender, MouseEventArgs e)
        {
            if (EmailType2Button.Background != new SolidColorBrush(Color.FromRgb(100, 150, 255)))
            {
                EmailType2Button.Background = new SolidColorBrush(Color.FromRgb(100, 150, 255));
                EmailType2Button.Foreground = Brushes.White;
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
        private void EmailType2Button_MouseLeave(object sender, MouseEventArgs e)
        {
            // Only reset if this button is not the currently selected one
            // Check if the content area shows EmailType1Page
            if (MainContentArea.Content is EmailType2Page)
            {
                // Keep selected state - do nothing
                return;
            }
            else
            {
                // Reset to default state
                EmailType2Button.Background = Brushes.White;
                EmailType2Button.Foreground = new SolidColorBrush(Color.FromRgb(39, 37, 55));
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