using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using Xceed.Document.NET;
using Xceed.Words.NET;
using Interop = Microsoft.Office.Interop.Word;
using OpenXml = DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXmlForVsto.Word.BenchmarkAddin {
    public partial class Ribbon1 {
        void Ribbon1_Load(object sender, RibbonUIEventArgs e) { }

        void ButtonWrite1RunVsto_Click(object sender, RibbonControlEventArgs e) => WriteVsto(1);

        void ButtonWrite100RunsVsto_Click(object sender, RibbonControlEventArgs e) => WriteVsto(10);

        void ButtonWrite10kRunsVsto_Click(object sender, RibbonControlEventArgs e) => WriteVsto(100);

        void ButtonWrite100kRunsVsto_Click(object sender, RibbonControlEventArgs e) => WriteVsto(316);

        void ButtonReadVsto_Click(object sender, RibbonControlEventArgs e) => ReadVsto();

        async void ButtonWrite1RunOpenXML_Click(object sender, RibbonControlEventArgs e) => await WriteOpenXml(1);

        async void ButtonWrite100RunsOpenXML_Click(object sender, RibbonControlEventArgs e) => await WriteOpenXml(10);

        async void ButtonWrite10kRunsOpenXML_Click(object sender, RibbonControlEventArgs e) => await WriteOpenXml(100);

        async void ButtonWrite1mRunsOpenXML_Click(object sender, RibbonControlEventArgs e) => await WriteOpenXml(1000);

        async void ButtonWrite100kRunsOpenXML_Click(object sender, RibbonControlEventArgs e) => await WriteOpenXml(316);

        async void ButtonReadOpenXML_Click(object sender, RibbonControlEventArgs e) => await ReadOpenXml();

        Interop.Application _app => Globals.ThisAddIn.Application;
        readonly Random _rand = new Random();

        void WriteVsto(int size) {
            Stopwatch sw = Stopwatch.StartNew();

            _app.ScreenUpdating = false;

            Interop.WdColor[] colors = {
                Interop.WdColor.wdColorRed,
                Interop.WdColor.wdColorGreen,
                Interop.WdColor.wdColorBlue,
                Interop.WdColor.wdColorGold,
                Interop.WdColor.wdColorAqua,
                Interop.WdColor.wdColorPink,
                Interop.WdColor.wdColorDarkBlue };
            Interop.Document document = _app.ActiveDocument;
            for (int i = 1; i <= size; i++) {
                Interop.Paragraph paragraph = document.Range().Paragraphs.Add();
                for (int j = 1; j <= size; j++) {
                    Interop.Range range = paragraph.Range;
                    range.Collapse(Interop.WdCollapseDirection.wdCollapseEnd);
                    range.Text = " " + (1000 * i + j);
                    Interop.Font font = range.Font;
                    font.Fill.BackColor.RGB = GenerateRandomColor();
                    font.Color = colors[_rand.Next(6)];
                    font.Size = _rand.Next(9, 14);
                    font.Italic = -_rand.Next(2);
                    font.Bold = -_rand.Next(2);
                }
            }

            _app.ScreenUpdating = true;

            sw.Stop();
            MessageBox.Show(sw.Elapsed.ToString(), $"VSTO: write {size * size}");
        }

        void ReadVsto() {
            Stopwatch sw = Stopwatch.StartNew();

            _app.ScreenUpdating = false;

            Interop.Document document = _app.ActiveDocument;
            int i = 0;
            foreach (Interop.Range range in document.Range().Words) {
                string a1 = range.Text;
                Interop.Font font = range.Font;
                int a2 = font.Fill.BackColor.RGB;
                Interop.WdColor a3 = font.Color;
                float a4 = font.Size;
                int a5 = font.Italic;
                int a6 = font.Bold;
                i++;
            }

            _app.ScreenUpdating = true;

            sw.Stop();
            MessageBox.Show(sw.Elapsed.ToString(), $"VSTO: read {i}");
        }

        async Task WriteOpenXml(int size) {
            Stopwatch sw = Stopwatch.StartNew();

            _app.ScreenUpdating = false;

            Interop.Range range = _app.ActiveDocument.Range();
            string tmpFile = new OpenXmlHelper().CopyToFile(range);

            _app.ScreenUpdating = true;

            await Task.Run(() => {
                Color[] colors = { Color.Red, Color.Green, Color.Blue, Color.Gold, Color.Gray, Color.Pink, Color.GreenYellow };
                using (DocX doc = DocX.Load(tmpFile)) {
                    for (int i = 1; i <= size; i++) {
                        Paragraph paragraph = doc.InsertParagraph();
                        for (int j = 1; j <= size; j++) {
                            Paragraph p = paragraph.Append(" " + 1000 * i + j)
                                .Shading(Color.FromArgb(GenerateRandomColor()))
                                .Color(colors[_rand.Next(6)])
                                .FontSize(_rand.Next(9, 14));
                            if (_rand.Next(2) == 1) {
                                p.Bold();
                            }
                            if (_rand.Next(2) == 1) {
                                p.Italic();
                            }
                        }
                    }

                    doc.Save();
                }
            });

            range.Text = " ";
            _app.ScreenUpdating = false;
            new OpenXmlHelper().CopyFromFile(tmpFile, range);
            _app.ScreenUpdating = true;

            File.Delete(tmpFile);

            sw.Stop();
            MessageBox.Show(sw.Elapsed.ToString(), $"OpenXML: write {size * size}");
        }

        async Task ReadOpenXml() {
            Stopwatch sw = Stopwatch.StartNew();

            _app.ScreenUpdating = false;

            Interop.Range range = _app.ActiveDocument.Range();

            string tmpFile = new OpenXmlHelper().CopyToFile(range);

            _app.ScreenUpdating = true;

            int i = 0;
            await Task.Run(() => {
                using (WordprocessingDocument doc = WordprocessingDocument.Open(tmpFile, true)) {
                    IEnumerable<OpenXml.Paragraph> paragraphs = doc.MainDocumentPart
                        .Document
                        .Body
                        .Elements<OpenXml.Paragraph>();
                    foreach (OpenXml.Paragraph paragraph in paragraphs) {
                        foreach (OpenXml.Run run in paragraph.Elements<OpenXml.Run>()) {
                            OpenXml.RunProperties properties = run.RunProperties;
                            string text = run.Elements<OpenXml.Text>().FirstOrDefault()?.Text;
                            OpenXml.Bold bold = properties?.Bold;
                            OpenXml.Italic italic = properties?.Italic;
                            OpenXml.Color color = properties?.Color;
                            StringValue backColor = properties?.Shading.Fill;
                            OpenXml.FontSize fontSize = properties?.FontSize;
                            i++;
                        }
                    }
                }
            });

            File.Delete(tmpFile);

            sw.Stop();
            MessageBox.Show(sw.Elapsed.ToString(), $"OpenXML: read {i}");
        }

        int GenerateRandomColor() =>
            _rand.Next(255) * 256 * 256 + _rand.Next(255) * 256 + _rand.Next(255);
    }
}