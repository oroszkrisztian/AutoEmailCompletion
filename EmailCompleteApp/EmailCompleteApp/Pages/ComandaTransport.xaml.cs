using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using EmailCompleteApp.ViewModels;

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

        private void ComboBox_GotFocus(object sender, RoutedEventArgs e)
        {
            if (sender is ComboBox combo)
            {
                combo.IsDropDownOpen = true;
                
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
                            textBox.GotFocus += (s, args) => combo.IsDropDownOpen = true;
                            textBox.MouseDown += (s, args) => combo.IsDropDownOpen = true;
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
            if (sender is ComboBox combo && combo.IsEditable && combo.SelectedItem != null)
            {
                combo.Text = combo.SelectedItem.ToString();
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
    }
}