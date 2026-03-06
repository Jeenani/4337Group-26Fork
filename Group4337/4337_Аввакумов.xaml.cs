using ClosedXML.Excel;
using Group4337.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace Group4337
{
    /// <summary>
    /// Логика взаимодействия для _4337_Аввакумов.xaml
    /// </summary>
    public partial class _4337_Аввакумов : Window
    {
        public _4337_Аввакумов()
        {
            InitializeComponent();

        }
        private void ImportExcel(string path)
        {
            using var ctx = new Isrpo3labContext();
            using var workbook = new XLWorkbook(path);
            var sheet = workbook.Worksheet(1);
            var rows = sheet.RangeUsed().RowsUsed().Skip(1);
            foreach (var row in rows)
            {
                var service = new Service
                {
                    ServiceName = row.Cell(2).GetString(),
                    ServiceType = row.Cell(3).GetString(),
                    Price = row.Cell(4).GetValue<decimal>()
                };
                ctx.Services.Add(service);
            }
            ctx.SaveChanges();
        }
        private void WriteSheet(XLWorkbook workbook, string name, List<Service> data)
        {
            var ws = workbook.Worksheets.Add(name);
            ws.Cell(1, 1).Value = "Id";
            ws.Cell(1, 2).Value = "Название";
            ws.Cell(1, 3).Value = "Тип";
            ws.Cell(1, 4).Value = "Цена";

            int row = 2;

            foreach (var s in data)
            {
                ws.Cell(row, 1).Value = s.Id;
                ws.Cell(row, 2).Value = s.ServiceName;
                ws.Cell(row, 3).Value = s.ServiceType;
                ws.Cell(row, 4).Value = s.Price;
                row++;
            }
        }
        private void ExportExcel(string path)
        {
            using var ctx = new Isrpo3labContext();
            var services = ctx.Services.ToList();
            var cat1 = services.Where(s => s.Price <= 350).ToList();
            var cat2 = services.Where(s => s.Price > 250 && s.Price <= 800).ToList();
            var cat3 = services.Where(s => s.Price > 800).ToList();
            using var wb = new XLWorkbook();
            WriteSheet(wb, "Категория 1", cat1);
            WriteSheet(wb, "Категория 2", cat2);
            WriteSheet(wb, "Категория 3", cat3);
            wb.SaveAs(path);
        }

        private void Import_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new Microsoft.Win32.OpenFileDialog();
            dialog.Filter = "Excel (*xlsx)|*.xlsx";
            if (dialog.ShowDialog() == true)
            {
                ImportExcel(dialog.FileName);
                MessageBox.Show("Данные успешно импортированы!", "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private void Export_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new Microsoft.Win32.SaveFileDialog();
            dialog.Filter = "Excel (*xlsx)|*.xlsx";
            if(dialog.ShowDialog() == true)
            {
                ExportExcel(dialog.FileName);
                MessageBox.Show("Экспорт завершен");
            }
        }

    }
}
