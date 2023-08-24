using Avalonia.Animation;
using Avalonia.Controls;
using Avalonia.Interactivity;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Information;
using ReactiveUI;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Security.Cryptography;
using System.Text;
using System.Windows.Input;

namespace Operations.ViewModels
{
    public class MainWindowViewModel : ViewModelBase
    {
        public string Greeting => "Welcome to Avalonia!";

        public List<TypeOperationViewModel> TypeOperations { get; set; }

        public ICommand CalculationCommand { get; set; }

        public object? Pk { get; set; }

        public object? Pob { get; set; }

        public object? Pi { get; set; }

        public double? Pop { get; set; }

        public double? Pisp { get; set; }

        public double? Pd { get; set; }

        object? _r;

        public object? R
        {
            get
            {
                return _r;
            }
            set
            {
                if (value == _r)
                    return;
                _r = value;
                try
                {
                    this.RaisePropertyChanged(nameof(R));
                    if (Convert.ToInt32(R) > 0)
                    {
                        TypeOperations = new List<TypeOperationViewModel>();
                        for (int i = 1; i <= Convert.ToInt32(R); i++)
                        {
                            TypeOperations.Add(new TypeOperationViewModel("Операция " + i.ToString(), i));
                        }
                        this.RaisePropertyChanged(nameof(TypeOperations));
                    }
                }
                catch
                {
                }
            }
        }

        public MainWindowViewModel()
        {
            TypeOperations = new List<TypeOperationViewModel>();
            CalculationCommand = ReactiveCommand.Create(Calculation);
        }

        public void Download(string fileName)
        {
            try
            {
                var newFile = new FileInfo(fileName);
                var Excel_Package = new ExcelPackage(newFile);
                var workSheet = Excel_Package.Workbook.Worksheets[0];
                var cells = workSheet.Cells;
                foreach (var type in TypeOperations)
                {
                    type.AddData(cells);
                }
                Pk = Convert.ToDouble(cells["F3"].Value.ToString().Replace(".", ","));
                this.RaisePropertyChanged(nameof(Pk));
                Pob = Convert.ToDouble(cells["G3"].Value.ToString().Replace(".", ","));
                this.RaisePropertyChanged(nameof(Pob));
                Pi = Convert.ToDouble(cells["H3"].Value.ToString().Replace(".", ","));
                this.RaisePropertyChanged(nameof(Pi));
                this.RaisePropertyChanged(nameof(TypeOperations));

            }
            catch
            {
            }
        }

        public void Calculation()
        {
            try
            {
                TypeOperations.ForEach(o => o.Calculation());
                Pop = 1;
                foreach (var oper in TypeOperations)
                {
                    Pop = Pop * Math.Pow(Convert.ToDouble(oper.P), Convert.ToDouble(oper.k));
                }
                Pisp = Convert.ToDouble(Pk) * Convert.ToDouble(Pob) * Convert.ToDouble(Pi);
                Pd = Pop + (1 - Pop) * Pisp;
                this.RaisePropertyChanged(nameof(Pop));
                this.RaisePropertyChanged(nameof(Pisp));
                this.RaisePropertyChanged(nameof(Pd));
            }
            catch
            {

            }
        }

    }
}
