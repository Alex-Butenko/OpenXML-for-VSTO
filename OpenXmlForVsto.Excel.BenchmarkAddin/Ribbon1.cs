using ClosedXML.Excel;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Diagnostics;
using System.Drawing;
using System.Threading.Tasks;
using System.Windows.Forms;
using Application = Microsoft.Office.Interop.Excel.Application;
using Font = Microsoft.Office.Interop.Excel.Font;

namespace OpenXmlForVsto.Excel.BenchmarkAddin {
    public partial class Ribbon1 {
        void Ribbon1_Load(object sender, RibbonUIEventArgs e) { }

        void ButtonWrite1CellVsto_Click(object sender, RibbonControlEventArgs e) => WriteVsto(1);

        void ButtonWrite100CellsVsto_Click(object sender, RibbonControlEventArgs e) => WriteVsto(10);

        void ButtonWrite10kCellsVsto_Click(object sender, RibbonControlEventArgs e) => WriteVsto(100);

        void ButtonWrite1mCellsVsto_Click(object sender, RibbonControlEventArgs e) => WriteVsto(1000);

        void ButtonRead1CellVsto_Click(object sender, RibbonControlEventArgs e) => ReadVsto(1);

        void ButtonRead100CellsVsto_Click(object sender, RibbonControlEventArgs e) => ReadVsto(10);

        void ButtonRead10kCellsVsto_Click(object sender, RibbonControlEventArgs e) => ReadVsto(100);

        void ButtonRead1mCellsVsto_Click(object sender, RibbonControlEventArgs e) => ReadVsto(1000);

        async void ButtonWrite1CellOpenXML_Click(object sender, RibbonControlEventArgs e) => await WriteOpenXml(1);

        async void ButtonWrite100CellsOpenXML_Click(object sender, RibbonControlEventArgs e) => await WriteOpenXml(10);

        async void ButtonWrite10kCellsOpenXML_Click(object sender, RibbonControlEventArgs e) => await WriteOpenXml(100);

        async void ButtonWrite1mCellsOpenXML_Click(object sender, RibbonControlEventArgs e) => await WriteOpenXml(1000);

        async void ButtonWrite9mCellsOpenXML_Click(object sender, RibbonControlEventArgs e) => await WriteOpenXml(3000);

        async void ButtonRead1CellOpenXML_Click(object sender, RibbonControlEventArgs e) => await ReadOpenXml(1);

        async void ButtonRead100CellsOpenXML_Click(object sender, RibbonControlEventArgs e) => await ReadOpenXml(10);

        async void ButtonRead10kCellsOpenXML_Click(object sender, RibbonControlEventArgs e) => await ReadOpenXml(100);

        async void ButtonRead1mCellsOpenXML_Click(object sender, RibbonControlEventArgs e) => await ReadOpenXml(1000);

        async void ButtonRead9mCellsOpenXML_Click(object sender, RibbonControlEventArgs e) => await ReadOpenXml(3000);

        Application _app => Globals.ThisAddIn.Application;
        readonly Random _rand = new Random();

        void WriteVsto(int size) {
            Stopwatch sw = Stopwatch.StartNew();

            _app.ScreenUpdating = false;

            Color[] colors = { Color.Red, Color.Green, Color.Blue,
                        Color.Bisque, Color.Gray, Color.Pink, Color.GreenYellow };
            string[] numberFormats = { "0.00", "0%", "£#,##0;-£#,##0", "#,##0;[Red]-#,##0" };
            Worksheet sheet = _app.ActiveSheet;
            for (int i = 1; i <= size; i++) {
                for (int j = 1; j <= size; j++) {
                    Range cell = sheet.Cells[i, j];
                    cell.Value = 1000 * i + j;
                    cell.Interior.Color = colors[_rand.Next(6)];
                    Font font = cell.Font;
                    font.Color = colors[_rand.Next(6)];
                    font.Size = _rand.Next(9, 14);
                    font.Italic = _rand.Next(1) == 1;
                    font.Bold = _rand.Next(1) == 1;
                    cell.NumberFormat = numberFormats[_rand.Next(3)];
                }
            }

            _app.ScreenUpdating = true;

            sw.Stop();
            MessageBox.Show(sw.Elapsed.ToString(), $"VSTO: write {size * size}");
        }

        void ReadVsto(int size) {
            Stopwatch sw = Stopwatch.StartNew();

            _app.ScreenUpdating = false;

            Worksheet sheet = _app.ActiveSheet;
            for (int i = 1; i <= size; i++) {
                for (int j = 1; j <= size; j++) {
                    Range cell = sheet.Cells[i, j];
                    dynamic a1 = cell.Value;
                    dynamic a2 = cell.Interior.Color;
                    Font font = cell.Font;
                    dynamic a3 = font.Color;
                    dynamic a4 = font.Size;
                    dynamic a5 = font.Italic;
                    dynamic a6 = font.Bold;
                    dynamic a7 = cell.NumberFormat;
                }
            }

            _app.ScreenUpdating = true;

            sw.Stop();
            MessageBox.Show(sw.Elapsed.ToString(), $"VSTO: read {size * size}");
        }

        async Task WriteOpenXml(int size) {
            Stopwatch sw = Stopwatch.StartNew();

            OpenXmlHelper oxh = new OpenXmlHelper();

            Worksheet sheet = _app.ActiveSheet;
            Range range = sheet.Range[sheet.Cells[1, 1], sheet.Cells[size, size]];

            _app.ScreenUpdating = false;
            string file = oxh.CopyToFile(range);
            _app.ScreenUpdating = true;

            await Task.Run(() => {
                XLColor[] colors = {
                XLColor.Red, XLColor.Green, XLColor.Blue,
                XLColor.BlueGray, XLColor.Gray, XLColor.Pink, XLColor.GreenYellow };
                string[] numberFormats = { "0.00", "0%", "£#,##0;-£#,##0", "#,##0;[Red]-#,##0" };
                Random rand = new Random();
                using (XLWorkbook workbook = new XLWorkbook(file)) {
                    IXLWorksheet worksheet = workbook.Worksheets.Worksheet(1);
                    for (int i = 1; i <= size; i++) {
                        for (int j = 1; j <= size; j++) {
                            IXLCell cell = worksheet.Cell(i, j);
                            int value = size * i + j;
                            cell.Value = value;
                            cell.Style.Fill.SetBackgroundColor(colors[rand.Next(6)]);
                            cell.Style.Font.SetFontColor(colors[rand.Next(6)]);
                            cell.Style.Font.SetFontSize(rand.Next(9, 14));
                            cell.Style.Font.Italic = rand.Next(1) == 1;
                            cell.Style.Font.Bold = rand.Next(1) == 1;
                            cell.Style.NumberFormat.SetFormat(numberFormats[rand.Next(3)]);
                        }
                    }
                    workbook.Save();
                }
            });

            _app.ScreenUpdating = false;
            oxh.CopyFromFile(file, range);
            _app.ScreenUpdating = true;

            sw.Stop();
            MessageBox.Show(sw.Elapsed.ToString(), $"OpenXML: write {size * size}");
        }

        async Task ReadOpenXml(int size) {
            Stopwatch sw = Stopwatch.StartNew();

            OpenXmlHelper oxh = new OpenXmlHelper();

            Worksheet sheet = _app.ActiveSheet;
            Range range = sheet.Range[sheet.Cells[1, 1], sheet.Cells[size, size]];

            _app.ScreenUpdating = false;
            string file = oxh.CopyToFile(range);
            _app.ScreenUpdating = true;

            await Task.Run(() => {
                using (XLWorkbook workbook = new XLWorkbook(file)) {
                    IXLWorksheet worksheet = workbook.Worksheets.Worksheet(1);
                    for (int i = 1; i <= size; i++) {
                        for (int j = 1; j <= size; j++) {
                            IXLCell cell = worksheet.Cell(i, j);
                            var a1 = cell.Value;
                            var a2 = cell.Style.Fill.BackgroundColor.Color;
                            var a3 = cell.Style.Font.FontColor.Color;
                            var a4 = cell.Style.Font.FontSize;
                            var a5 = cell.Style.Font.Italic;
                            var a6 = cell.Style.Font.Bold;
                            var a7 = cell.Style.NumberFormat.Format;
                        }
                    }
                }
            });

            _app.ScreenUpdating = false;
            oxh.CopyFromFile(file, range);
            _app.ScreenUpdating = true;

            sw.Stop();
            MessageBox.Show(sw.Elapsed.ToString(), $"OpenXML: read {size * size}");
        }
    }
}