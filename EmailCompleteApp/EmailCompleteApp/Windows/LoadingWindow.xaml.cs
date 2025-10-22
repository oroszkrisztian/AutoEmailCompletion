using System.Windows;

namespace EmailCompleteApp.Windows
{
    public partial class LoadingWindow : Window
    {
        public LoadingWindow()
        {
            InitializeComponent();
        }

        public void UpdateProgress(string message)
        {
            Dispatcher.Invoke(() =>
            {
                LoadingText.Text = message;
            });
        }

        public void UpdateDetail(string detail)
        {
            Dispatcher.Invoke(() =>
            {
                ProgressText.Text = detail;
            });
        }
    }
}