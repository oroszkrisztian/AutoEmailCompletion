using CommunityToolkit.Mvvm.ComponentModel;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows.Data;

namespace EmailCompleteApp.ViewModels
{
    public partial class MainViewModel : ObservableObject
    {
        [ObservableProperty]
        private ObservableCollection<string> allItems;

        [ObservableProperty]
        private string filterText;

        private CollectionViewSource filteredItemsViewSource;

        public ICollectionView FilteredItems => filteredItemsViewSource.View;

        public MainViewModel()
        {
            // Initialize with sample fruit data
            AllItems = new ObservableCollection<string>
            {
                "Apple",
                "Apricot",
                "Banana",
                "Blueberry",
                "Cherry",
                "Grape",
                "Grapefruit",
                "Kiwi",
                "Lemon",
                "Lime",
                "Mango",
                "Orange",
                "Papaya",
                "Peach",
                "Pear",
                "Pineapple",
                "Plum",
                "Raspberry",
                "Strawberry",
                "Watermelon"
            };

            // Setup CollectionViewSource for filtering
            filteredItemsViewSource = new CollectionViewSource
            {
                Source = AllItems
            };

            // Set filter predicate
            filteredItemsViewSource.Filter += FilterItems;
        }

        partial void OnFilterTextChanged(string value)
        {
            // Refresh the filter when FilterText changes
            filteredItemsViewSource.View.Refresh();
        }

        private void FilterItems(object sender, FilterEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(FilterText))
            {
                e.Accepted = true;
            }
            else
            {
                var item = e.Item as string;
                e.Accepted = item != null && item.Contains(FilterText, StringComparison.OrdinalIgnoreCase);
            }
        }
    }
}
