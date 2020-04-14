using System;
using System.Collections.Generic;
using System.Windows;
using System.Net;
using System.Windows.Controls;
using System.IO;
using OfficeOpenXml;
using System.Windows.Input;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Text;

namespace lab2
{
    public partial class MainWindow : Window
    {
        public string d = @"D:\test2.xlsx";
        public string defaultD = null;
        public MainWindow()
        {
            InitializeComponent();
        }
        private void Download (object sender, RoutedEventArgs R)
        {
            if (System.IO.File.Exists(d)) { System.Windows.MessageBox.Show("Файл уже существует в данном каталоге. Файл не был скачан."); return; }
                WebClient Client = new WebClient();
            Client.DownloadFile("https://bdu.fstec.ru/files/documents/thrlist.xlsx", d);
            System.Windows.MessageBox.Show("Скачивание прошло успешно! Файл хранится здесь: " + d);
            defaultD = d;
        }
        private void Parse (object sender, RoutedEventArgs R) 
        {
            //List<Threat> threats = new List<Threat>();
            if (!System.IO.File.Exists(d)) { System.Windows.MessageBox.Show("Файла не существует в данном каталоге. Парсинг не был осуществлён."); return; }
            FileInfo fi = new FileInfo(d);
            Threat.threats.Clear();
            using (ExcelPackage excelPackage = new ExcelPackage(fi))
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets[0];
                for (int i = worksheet.Dimension.Start.Row + 2; i <= worksheet.Dimension.End.Row; i++)
                {
                    Threat threat = new Threat();
                    threat.Id = int.Parse(worksheet.Cells[i, 1].Value.ToString());
                    threat.Name = worksheet.Cells[i, 2].Value.ToString();
                    threat.Notice = worksheet.Cells[i, 3].Value.ToString();
                    threat.Source = worksheet.Cells[i, 4].Value.ToString();
                    threat.Influence = worksheet.Cells[i, 5].Value.ToString();
                    if (worksheet.Cells[i, 6].Value.ToString().Equals("1")) { threat.ConfidentityThreat = true; }
                    if (worksheet.Cells[i, 7].Value.ToString().Equals("1")) { threat.IntegrityThreat = true; }
                    if (worksheet.Cells[i, 8].Value.ToString().Equals("1")) { threat.AccessThreat = true; }
                    DateTime dt = DateTime.FromOADate((double)worksheet.Cells[i, 9].Value);
                    threat.CreationDate = dt.ToShortDateString ();
                    dt = DateTime.FromOADate((double)worksheet.Cells[i, 10].Value);
                    threat.ChangeDate = dt.ToShortDateString();
                    Threat.threats.Add(threat);
                }
                ListView.ItemsSource = Threat.threats;
            } 
        }
        private void Refresh (object sender, RoutedEventArgs R)
        {
            if (!System.IO.File.Exists(d)) { System.Windows.MessageBox.Show("Файл не существует в данном каталоге. Файл не был обновлён."); return; }
            FileInfo fi = new FileInfo(d);

            using (ExcelPackage excelPackage = new ExcelPackage(fi))
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets[0];
                int i=3;
                
                string s;
                //DateTime dt;
                bool b = false;
                foreach (Threat t in Threat.threats)
                {
                    t.Changes = "\n";
                    s = worksheet.Cells[i, 2].Value.ToString();
                    if (!t.Name.Equals(s)) 
                    { 
                        t.Changes = "Было измененно Имя. Было: "+t.Name+ "\nСтало: " + s; t.Name = s+"\n"; 
                    }
                    s = worksheet.Cells[i, 3].Value.ToString();
                    if (!t.Notice.Equals(s))
                    {
                        t.Changes = t.Changes + "Было измененно Описание. Было: " + t.Notice + "\nСтало: " + s; t.Notice = s + "\n";
                    }
                    s = worksheet.Cells[i, 4].Value.ToString();
                    if (!t.Source.Equals(s))
                    {
                        t.Changes =t.Changes + "Был изменен Источник угрозы. Было: " + t.Source + "\nСтало: " + s; t.Source = s + "\n";
                    }
                    s = worksheet.Cells[i, 5].Value.ToString();
                    if (!t.Influence.Equals(s))
                    {
                        t.Changes = t.Changes + "Был изменен Объект воздействия. Было: " + t.Influence + "\nСтало: " + s + "\n";
                    }
                    if (worksheet.Cells[i, 6].Value.ToString().Equals("1")) { b = true; }
                    if (worksheet.Cells[i, 6].Value.ToString().Equals("0")) { b = false; }

                    if (!(b == t.ConfidentityThreat)) 
                    {
                        if (b) { t.Changes = t.Changes + "Появилось последствие Нарушение конфинденциальности" + "\n"; } else { t.Changes = t.Changes + "Исчезло последствие Нарушение конфинденциальности" + "\n"; }
                    }
                    if (worksheet.Cells[i, 7].Value.ToString().Equals("1")) { b = true; }
                    if (worksheet.Cells[i, 7].Value.ToString().Equals("0")) { b = false; }
                    if (!(b == t.IntegrityThreat))
                    {
                        if (b) { t.Changes = t.Changes + "Появилось последствие Нарушение целостности" + "\n"; } else { t.Changes = t.Changes + "Исчезло последствие Нарушение целостности" + "\n"; }
                    }
                    if (worksheet.Cells[i, 8].Value.ToString().Equals("1")) { b = true; }
                    if (worksheet.Cells[i, 8].Value.ToString().Equals("0")) { b = false; }
                    if (!(b == t.AccessThreat))
                    {
                        if (b) { t.Changes = t.Changes + "Появилось последствие Нарушение доступности" + "\n"; } else { t.Changes = t.Changes + "Исчезло последствие Нарушение доступности" + "\n"; }
                    }
                    i++;
                    if (t.Changes.Equals("\n")) { t.Changes = "-"; }
                    b = false;
                    /*dt = DateTime.FromOADate((double)worksheet.Cells[i, 9].Value);
                    s = dt.ToShortDateString();
                    if (!t.CreationDate.Equals(s)) { t.Changes = t.Changes + "Изменилась Дата включения. Было: " + t.CreationDate + "Стало: " + s + "\n";  }
                    dt = DateTime.FromOADate((double)worksheet.Cells[i, 10].Value);
                    s = dt.ToShortDateString();
                    if (!t.ChangeDate.Equals(s)) { t.Changes = t.Changes + "Изменилась Дата последнего изменения. Было: " + t.ChangeDate + "Стало: " + s + "\n"; } */
                }
                i = 0;
                foreach (Threat t in Threat.threats)
                {
                    if (!t.Changes.Equals("-"))
                    {
                        i++;
                        System.Windows.MessageBox.Show("Для угроза с ID - " + t.Id + "\n" + t.Changes + "\nДата изменения: " +t.ChangeDate);
                    }
                }
                if (i != 0) { System.Windows.MessageBox.Show("Обновление прошло успешно, всего обновленных записей: " + i.ToString()); } else { System.Windows.MessageBox.Show("Список записей не изменился"); }
                
            }
        }
        public void Save (object sender, RoutedEventArgs R)
        {

            SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(d, SpreadsheetDocumentType.Workbook);
            WorkbookPart workbookpart = spreadsheetDocument.AddWorkbookPart();
            workbookpart.Workbook = new Workbook();
            WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());

            Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.
                AppendChild<Sheets>(new Sheets());
            Sheet sheet = new Sheet()
            {
                Id = spreadsheetDocument.WorkbookPart.
                GetIdOfPart(worksheetPart),
                SheetId = 1,
                Name = "Лист 1"

            };sheets.Append(sheet);


            workbookpart.Workbook.Save();
            spreadsheetDocument.Close();
            /*int i = 1;
            using (ExcelPackage excelPackage = new ExcelPackage())
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets[0];
                foreach (Threat threat in Threat.threats)
                {
                    worksheet.Cells[i, 1] = threat.Id;
                    worksheet.Cells[i, 1].Value = threat.Id;
                    worksheet.Cells[i, 2].Value = threat.Name;
                    worksheet.Cells[i, 3].Value = threat.Notice;
                    worksheet.Cells[i, 4].Value = threat.Source;
                    worksheet.Cells[i, 5].Value = threat.Influence;

                    if (threat.ConfidentityThreat) { worksheet.Cells[i, 6].Value = 1; } else { worksheet.Cells[i, 6].Value = 0; }
                    if (threat.IntegrityThreat) { worksheet.Cells[i, 7].Value = 1; } else { worksheet.Cells[i, 7].Value = 0; }
                    if (threat.AccessThreat) { worksheet.Cells[i, 8].Value = 1; } else { worksheet.Cells[i, 8].Value = 0; }

                    worksheet.Cells[i, 9].Value = threat.CreationDate.ToString();
                    worksheet.Cells[i, 10].Value = threat.ChangeDate.ToString();
                    i++;
                }
            }*/



            System.Windows.MessageBox.Show("Файл сохранён здесь: " + d);
        }
        private void ShortInfo(object sender, MouseButtonEventArgs a)
        {
            if (ListView.SelectedItem == null)
            {
                return;
            }
            Threat threat = ListView.SelectedItem as Threat;
            string s = "ID: " + threat.Id + "\nИмя: " + threat.Name + "\n\nОписание: " + threat.Notice + "\n\nИсточник угрозы: " + threat.Source + "\nОбъект воздействия: " + threat.Influence;
            if (threat.ConfidentityThreat) { s = s + "\nНарушение конфиденциальности: Да"; } else { s = s + "\nНарушение конфиденциальности: Нет"; }
            if (threat.IntegrityThreat) { s = s + "\nНарушение целостности: Да"; } else { s = s + "\nНарушение целостности: Нет"; }
            if (threat.AccessThreat) { s = s + "\nНарушение доступности: Да"; } else { s = s + "\nНарушение доступности: Нет"; }
            s = s + " \nДата включения:                      " + threat.CreationDate + " \nДата последнего изменения: " + threat.ChangeDate ;
            System.Windows.MessageBox.Show(s);
        }

        private void TextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            d = TextBox1.Text;
        }
    }
}
