using System;
using System.Windows;
using System.Windows.Controls;
using EmailCompleteApp.Models;
using EmailCompleteApp.ViewModels;

namespace EmailCompleteApp.Pages
{
    public partial class LocationPage : UserControl
    {
        private LocationViewModel? ViewModel => DataContext as LocationViewModel;

        public LocationPage()
        {
            InitializeComponent();
            DataContext = new LocationViewModel();
        }

        private void LocationDataGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (sender is DataGrid grid && grid.SelectedItem is Location location)
            {
                ViewModel?.EditLocation(location);
                grid.SelectedItem = null; // Clear selection after opening edit
            }
        }
    }
}
