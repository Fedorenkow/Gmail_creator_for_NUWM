using System;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Text.Json;
using Microsoft.Office.Interop.Excel;
using System.IO;
using System.Reflection;
using OfficeOpenXml;
using System.Linq;
using System.Diagnostics;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace Gmail_Creator
{
    public partial class Form1 : Form
    {
        public static string GetRandomPassword(int length)
        {
            const string chars = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ";

            StringBuilder sb = new StringBuilder();
            Random rnd = new Random();

            for (int i = 0; i < length; i++)
            {
                int index = rnd.Next(chars.Length);
                sb.Append(chars[index]);
            }

            return sb.ToString();
        }
       
        public Form1()
        {
            InitializeComponent();
        }

        public static async Task<string> Translator(string lastName)
        {
            string apiUrl = $"https://api.mymemory.translated.net/get?q={lastName.ToLower()}&langpair=uk|en";

            using (var httpClient = new HttpClient())
            {
                var response = await httpClient.GetAsync(apiUrl);
                response.EnsureSuccessStatusCode();

                var responseString = await response.Content.ReadAsStringAsync();
                var responseData = JsonSerializer.Deserialize<ApiResponse>(responseString);

                return responseData.responseData.translatedText;
            }
        }

        public class ApiResponse
        {
            public ResponseData responseData { get; set; }
        }

        public class ResponseData
        {
            public string translatedText { get; set; }
        }

        private async void button3_Click(object sender, EventArgs e)
        {
            long phoneNumber = long.Parse(textBox1.Text);
            textBox1.Text= phoneNumber.ToString();

            string lastName = textBox5.Text;
            string name = textBox6.Text;
            string surName = textBox7.Text;
            string gmail = "";
            string translatedLastName = await Translator(lastName);
            string translatedName = await Translator(name);
            translatedName = translatedName.ToLower();
            
            gmail += $"{translatedLastName}.";
            gmail += $"{translatedName[0]}_{comboBox1.SelectedValue}2023";
            gmail += "@nuwm.edu.ua";
            textBox8.Text = gmail;
            textBox9.Text = GetRandomPassword(8);

            string folderPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);

            // Формуємо повний шлях до файлу
            string filePath = Path.Combine(folderPath, "data.xlsx");

            // Перевіряємо, чи існує файл, та якщо ні - створюємо новий
            if (!File.Exists(filePath))
            {
                using (var package = new ExcelPackage())
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Worksheet1");
                    worksheet.Cells[1, 1].Value = "Призвіще";
                    worksheet.Cells[1, 2].Value = "Ім'я";
                    worksheet.Cells[1, 3].Value = "По-батькові";
                    worksheet.Cells[1, 4].Value = "Номер телефону";
                    worksheet.Cells[1, 5].Value = "Корпоративна пошта";
                    worksheet.Cells[1, 6].Value = "Пароль";
                    package.SaveAs(new FileInfo(filePath));
                }
            }

            // Записуємо дані до файлу
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.First();
                int lastRow = worksheet.Dimension.End.Row;
                
                worksheet.Cells[lastRow + 1, 1].Value = lastName;
                worksheet.Cells[lastRow + 1, 2].Value = name;
                worksheet.Cells[lastRow + 1, 3].Value = surName;
                worksheet.Cells[lastRow + 1, 4].Value = phoneNumber;
                worksheet.Cells[lastRow + 1, 5].Value = gmail;
                worksheet.Cells[lastRow + 1, 6].Value = textBox9.Text;
                package.Save();
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Отримуємо шлях до файлу
            string folderPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            string filePath = Path.Combine(folderPath, "data.xlsx");

            // Відкриваємо папку з файлом Excel
            Process.Start("explorer.exe", $"/select,\"{filePath}\"");
        }

        
        private void Form1_Load(object sender, EventArgs e)
        {
            comboBox1.DisplayMember = "Text";
            comboBox1.ValueMember = "Value";
            var items = new[] {
                new { Text = "Автоматики, кібернетики та обчислювальної техніки", Value = "ak"},
                new { Text = "Водного господартсва та природооблаштування", Value = "vg" },
                new { Text = "Будівництва та Ахрітектури", Value = "ba" },
                new { Text = "Агроекології та землеустрою", Value = "az"},
                new { Text = "Механічний інститут", Value = "m" },
                new { Text = "Інститут права", Value = "p" },
                new {  Text = "Економіки та менеджменту", Value = "em" },
                new { Text = "Охорони здоров'я", Value = "oz"}
            };
            comboBox1.DataSource = items;
        }
    }
}
