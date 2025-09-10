using System.Windows;

namespace EmailCompleteApp.Pages
{
    public partial class LoadingWindow : Window
    {
        public LoadingWindow()
        {
            InitializeComponent();
            Topmost = true;
            ShowInTaskbar = false;
        }
    }
}


