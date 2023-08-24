using OfficeOpenXml;
using ReactiveUI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Operations.ViewModels
{
    public class TypeOperationViewModel : ViewModelBase
    {
        public string Name { get; set; }

        public int Number { get; set; }

        public int SelectedIndex { get; set; }

        public Dictionary<int, double?> Items => new Dictionary<int, double?>() { { 0, 0.0001 }, { 1, 0.001 }, { 2, 0.003 },
        { 3, 0.03 }, { 4, 0.2 }, { 5, 0.3 },{ 6, 0.01 },{ 7, 0.1 }};

        public object? T { get; set; }

        public object? Lyambda { get; set; }

        public object? N { get; set; }

        public object? n { get; set; }

        public object? P { get; set; }

        public object? k { get; set; }

        public TypeOperationViewModel(string name, int number)
        {
            Name = name;
            Number = number;
            Lyambda = 0.0001;
            SelectedIndex = 0;
        }


        public void AddData(ExcelRange cells)
        {
            Lyambda = Convert.ToDouble(cells["B" + (Number + 2).ToString()].Value.ToString().Replace(".", ","));
            if (Items.Any(o => o.Value == Convert.ToDouble(Lyambda)))
                SelectedIndex = Items.Where(o => o.Value == Convert.ToDouble(Lyambda)).First().Key;
            this.RaisePropertyChanged(nameof(SelectedIndex));
            N = Convert.ToDouble(cells["C" + (Number + 2).ToString()].Value.ToString().Replace(".", ","));
            T = Convert.ToDouble(cells["D" + (Number + 2).ToString()].Value.ToString().Replace(".", ","));
            k = Convert.ToDouble(cells["E" + (Number + 2).ToString()].Value.ToString().Replace(".", ","));
            this.RaisePropertyChanged(nameof(Lyambda));
            this.RaisePropertyChanged(nameof(N));
            this.RaisePropertyChanged(nameof(T));
            this.RaisePropertyChanged(nameof(k));
        }

        public void Calculation()
        {
            n = Convert.ToDouble(Lyambda) * Convert.ToDouble(N) * Convert.ToDouble(T);
            this.RaisePropertyChanged(nameof(n));
            P = 1 - Convert.ToDouble(n) / Convert.ToDouble(N);
            this.RaisePropertyChanged(nameof(P));
        }
    }
}
