using Avalonia.Controls;
using Avalonia.Interactivity;
using Operations.ViewModels;
using OfficeOpenXml.Style;
using OfficeOpenXml;
using System;
using System.ComponentModel;
using System.IO;
using System.Linq;

namespace Operations.Views
{
    public partial class MainWindow : Window
    {
        public MainWindowViewModel MainWindowViewModel => DataContext as MainWindowViewModel;

        public MainWindow()
        {
            InitializeComponent();
            //MainWindowViewModel = new MainWindowViewModel();
            //DataContext = MainWindowViewModel;
            //var comboBox = this.FindControl<ComboBox>("comboBox");
            //comboBox.DataContext = mainWindow;
        }

        public void SelectionChanged(object sender, SelectionChangedEventArgs args)
        {
            if ((sender as ComboBox).DataContext != null)
                ((sender as ComboBox).DataContext as TypeOperationViewModel).Lyambda = Convert.ToDouble(((args.AddedItems[0] as ComboBoxItem).Content as TextBlock).Text);
            //MainWindowViewModel.
            //var json = JsonConvert.SerializeObject(ViewModel.DtConfiguration);
            //File.WriteAllText("Models\\DtConfiguration.json", json);
        }

        public async void Download_Clicked(object sender, RoutedEventArgs args)
        {
            var fileDialog = new OpenFileDialog();
            fileDialog.Filters.Add(new FileDialogFilter() { Name = "Excel", Extensions = { "xlsx" } });
            var result = await fileDialog.ShowAsync(this);
            if (result != null)
            {
                (DataContext as MainWindowViewModel).Download(result.FirstOrDefault());
            }
        }

        public async void Export_Clicked(object sender, RoutedEventArgs args)
        {
            try
            {
                var fileDialog = new OpenFileDialog();
                fileDialog.Filters.Add(new FileDialogFilter() { Name = "Excel", Extensions = { "xlsx" } });
                var result = await fileDialog.ShowAsync(this);
                if (result != null)
                {
                    var newFile = new FileInfo(result.First());
                    var Excel_Package = new ExcelPackage(newFile);
                    var workSheet = Excel_Package.Workbook.Worksheets.First();
                    var cells = workSheet.Cells;
                    for (int i = 3; i < Convert.ToInt32(MainWindowViewModel.R) + 3; i++)
                    {
                        cells["I" + i.ToString()].Value = MainWindowViewModel.TypeOperations[i - 3].n;
                        cells["J" + i.ToString()].Value = MainWindowViewModel.TypeOperations[i - 3].P;

                    }
                    cells["K3"].Value = MainWindowViewModel.Pop;
                    cells["L3"].Value = MainWindowViewModel.Pisp;
                    cells["M3"].Value = MainWindowViewModel.Pd;
                    Excel_Package.Save();
                }
            }
            catch
            {

            }
        }
    }

}
