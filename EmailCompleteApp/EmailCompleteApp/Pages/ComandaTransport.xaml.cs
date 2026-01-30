using EmailCompleteApp.Models;
using EmailCompleteApp.ViewModels;
using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace EmailCompleteApp.Pages
{
    public partial class ComandaTransport : UserControl
    {
        private ComandaTransportViewModel? _viewModel;

        public ComandaTransport()
        {
            InitializeComponent();
            DataContext = new ComandaTransportViewModel();
            _viewModel = DataContext as ComandaTransportViewModel;
        }

        // Constructor for edit mode
        public ComandaTransport(HistoryTransport historyItem)
        {
            InitializeComponent();
            DataContext = new ComandaTransportViewModel(historyItem);
            _viewModel = DataContext as ComandaTransportViewModel;
        }

        private void ComboBox_GotFocus(object sender, RoutedEventArgs e)
        {
            if (sender is ComboBox combo)
            {
                // Do NOT open dropdown on focus; only when user types
                if (combo.Items.Count == 0 && string.IsNullOrWhiteSpace(combo.Text))
                {
                    if (_viewModel != null)
                    {
                        if (combo.Name == "ClientComboBox")
                        {
                            _viewModel.Client = "";
                        }
                        else if (combo.Name == "TransportatorComboBox")
                        {
                            _viewModel.Transportator = "";
                        }
                        else if (combo.Name == "IncarcareComboBox")
                        {
                            _viewModel.LocatieIncarcare = "";
                        }
                        else if (combo.Name == "DescarcareComboBox")
                        {
                            _viewModel.LocatieDescarcare = "";
                        }
                    }
                }
            }
        }

        private void ComboBox_Loaded(object sender, RoutedEventArgs e)
        {
            if (sender is ComboBox combo && combo.IsEditable)
            {
                combo.Dispatcher.BeginInvoke(() =>
                {
                    try
                    {
                        if (combo.Template.FindName("PART_EditableTextBox", combo) is TextBox textBox)
                        {
                            // Open dropdown ONLY when there are matching results
                            textBox.TextChanged += (s, args) =>
                            {
                                var hasText = !string.IsNullOrWhiteSpace(textBox.Text);
                                var hasItems = combo.Items.Count > 0;
                                combo.IsDropDownOpen = hasText && hasItems;
                            };
                        }
                    }
                    catch
                    {
                    }
                }, System.Windows.Threading.DispatcherPriority.Loaded);
            }
        }

        private void ComboBox_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (sender is ComboBox combo && combo.IsEditable)
            {
                if (e.Key == Key.Enter && combo.Items.Count > 0)
                {
                    combo.SelectedItem = combo.Items[0];
                    combo.IsDropDownOpen = false;
                    e.Handled = true;
                }
                else if (e.Key == Key.Escape)
                {
                    combo.IsDropDownOpen = false;
                    e.Handled = true;
                }
            }
        }

        private void ComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (sender is ComboBox combo && combo.IsEditable && combo.SelectedItem != null && _viewModel != null)
            {
                if(combo.SelectedItem is Location loc)
                {
                    if(combo.Name == "IncarcareComboBox")
                    {
                        _viewModel.SetUpdatingFromSelection(true);
                        _viewModel.UpdatePickupLocation(loc);
                        combo.Text = loc.ToString();
                        _viewModel.SetUpdatingFromSelection(false);
                    }
                    else if(combo.Name == "DescarcareComboBox")
                    {
                        _viewModel.SetUpdatingFromSelection(true);
                        _viewModel.UpdateDeliveryLocation(loc);
                        combo.Text = loc.ToString();
                        _viewModel.SetUpdatingFromSelection(false);
                    }
                } 
                else if(combo.SelectedItem is Transportator transportator)
                {
                    if (combo.Name == "TransportatorComboBox")
                    {
                        _viewModel.SetUpdatingFromSelection(true);
                        _viewModel.Transportator = transportator.ToString();
                        _viewModel.GetTermenPlata(transportator);
                        combo.Text = _viewModel.Transportator;
                        _viewModel.SetUpdatingFromSelection(false);
                    }
                }
                else if(combo.SelectedItem is Client client)
                {
                    if (combo.Name == "ClientComboBox")
                    {
                        _viewModel.SetUpdatingFromSelection(true);
                        _viewModel.Client = client.ToString();
                        combo.Text = _viewModel.Client;
                        _viewModel.SetUpdatingFromSelection(false);
                    }
                }
                else if(combo.SelectedItem is Product product)
                {
                    if (combo.Name == "ProdusComboBox")
                    {
                        _viewModel.SetUpdatingFromSelection(true);
                        _viewModel.Produs = product.ToString();
                        combo.Text = _viewModel.Produs;
                        _viewModel.SetUpdatingFromSelection(false);
                    }
                }
                else if(combo.SelectedItem is Contact contact)
                {
                    if (combo.Name == "ContactComboBox")
                    {
                        _viewModel.SetUpdatingFromSelection(true);
                        _viewModel.Contact = contact.ToString();
                        combo.Text = _viewModel.Contact;
                        _viewModel.SetUpdatingFromSelection(false);
                    }
                }
                combo.IsDropDownOpen = false;
            }
        }

        private void ComboBox_DropDownOpened(object sender, EventArgs e)
        {
            if (sender is ComboBox combo && combo.IsEditable)
            {
                combo.Dispatcher.BeginInvoke(() =>
                {
                    try
                    {
                        if (combo.Template.FindName("PART_EditableTextBox", combo) is TextBox tb)
                        {
                            var length = tb.Text?.Length ?? 0;
                            tb.SelectionStart = length;
                            tb.SelectionLength = 0;
                            tb.CaretIndex = length;
                            tb.Focus();
                        }
                    }
                    catch
                    {

                    }
                }, System.Windows.Threading.DispatcherPriority.Input);
            }
        }

        private void CommentTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            //autogrow height for textbox
            //need to scroll main page to bottom

            if (sender is TextBox textBox)
            {
                textBox.Height = Double.NaN;
                textBox.UpdateLayout();
                var desiredHeight = textBox.DesiredSize.Height;
                textBox.Height = desiredHeight;
                //find main scrollviewer
                Dispatcher.BeginInvoke(() =>
                {
                    var scrollViewer = FindAncestor<ScrollViewer>(this);
                    if (scrollViewer != null)
                    {
                        scrollViewer.ScrollToEnd();
                    }
                }, System.Windows.Threading.DispatcherPriority.Background);

            }
        }

        private static T? FindAncestor<T>(DependencyObject current) where T : DependencyObject
        {
            while (current != null)
            {
                if (current is T desired)
                {
                    return desired;
                }
                current = VisualTreeHelper.GetParent(current);
            }
            return null;
        }
    }
}