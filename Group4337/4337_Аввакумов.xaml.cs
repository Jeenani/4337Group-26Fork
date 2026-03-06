using ClosedXML.Excel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Group4337.Models;
using System.IO;
using System.Text.Json;
using System.Windows;

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
            if (dialog.ShowDialog() == true)
            {
                ExportExcel(dialog.FileName);
                MessageBox.Show("Экспорт завершен");
            }
        }

        private void ImportJson(string path)
        {
            using var ctx = new Isrpo3labContext();
            var json = File.ReadAllText(path);
            var options = new JsonSerializerOptions
            {
                PropertyNameCaseInsensitive = true
            };
            var services = JsonSerializer.Deserialize<List<Service>>(json, options);
            if (services == null) return;
            foreach (var s in services)
            {
                ctx.Services.Add(new Service
                {
                    ServiceName = s.ServiceName,
                    ServiceType = s.ServiceType,
                    Price = s.Price
                });
            }
            ctx.SaveChanges();
        }
        private void Import_Json_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "JSON (*.json)|*.json"
            };
            if (dialog.ShowDialog() == true)
            {
                ImportJson(dialog.FileName);
                MessageBox.Show("Данные успешно импортированы!", "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private TableCell MakeCell(string text, bool bold = false)
        {
            var run = bold
                ? new Run(new RunProperties(new Bold()), new Text(text))
                : new Run(new Text(text));
            return new TableCell(new Paragraph(run));
        }

        private void AddWordPage(Body body, string title, List<Service> data, bool addPageBreak)
        {
            body.AppendChild(new Paragraph(
                new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
                new Run(new RunProperties(new Bold()), new Text(title))
            ));

            var table = new Table();
            table.AppendChild(new TableProperties(
                new TableWidth { Width = "5000", Type = TableWidthUnitValues.Pct },
                new TableBorders(
                    new TopBorder { Val = BorderValues.Single, Size = 4 },
                    new BottomBorder { Val = BorderValues.Single, Size = 4 },
                    new LeftBorder { Val = BorderValues.Single, Size = 4 },
                    new RightBorder { Val = BorderValues.Single, Size = 4 },
                    new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4 },
                    new InsideVerticalBorder { Val = BorderValues.Single, Size = 4 }
                )
            ));

            var headerRow = new TableRow();
            foreach (var h in new[] { "Id", "Название", "Тип", "Цена" })
                headerRow.AppendChild(MakeCell(h, bold: true));
            table.AppendChild(headerRow);

            foreach (var s in data)
            {
                var row = new TableRow();
                foreach (var val in new[] { s.Id.ToString(), s.ServiceName, s.ServiceType, s.Price.ToString("F2") })
                    row.AppendChild(MakeCell(val));
                table.AppendChild(row);
            }

            body.AppendChild(table);

            if (addPageBreak)
                body.AppendChild(new Paragraph(new Run(new Break { Type = BreakValues.Page })));
        }
        private void ExportWord(string path)
        {
            using var ctx = new Isrpo3labContext();
            var services = ctx.Services.ToList();
            var cat1 = services.Where(s => s.Price <= 350).ToList();
            var cat2 = services.Where(s => s.Price > 250 && s.Price <= 800).ToList();
            var cat3 = services.Where(s => s.Price > 800).ToList();

            using var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document();
            var body = mainPart.Document.AppendChild(new Body());

            AddWordPage(body, "Категория 1 (0 — 350)", cat1, addPageBreak: true);
            AddWordPage(body, "Категория 2 (250 — 800)", cat2, addPageBreak: true);
            AddWordPage(body, "Категория 3 (от 800)", cat3, addPageBreak: false);

            mainPart.Document.Save();
        }
        private void Export_Word_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new Microsoft.Win32.SaveFileDialog { Filter = "Word (*.docx)|*.docx" };
            if (dialog.ShowDialog() == true)
            {
                ExportWord(dialog.FileName);
                MessageBox.Show("Экспорт в Word завершён!", "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }
    }
}
