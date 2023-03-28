using System;
using System.Text;
using System.Windows.Forms;
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
        public static string Transliterate(string input)
        {
            string[] ukr = { "А", "Б", "В", "Г","Ґ", "Д", "Е","Є","Ж", "З", "И","І", "Й",
      "К", "Л", "М", "Н", "О", "П", "Р", "С", "Т", "У", "Ф", "Х", "Ц",
      "Ч", "Ш", "Щ", "Ь", "Ю", "Я",
      "а", "б", "в", "г","ґ", "д", "е","є","ж", "з", "и","і", "й",
      "к", "л", "м", "н", "о", "п", "р", "с", "т", "у", "ф", "х", "ц",
      "ч", "ш", "щ", "ь", "ю", "я"};

            string[] eng = { "A", "B", "V", "H","G", "D", "E","Ye", "ZH", "Z", "I","I", "Y",
      "K", "L", "M", "N", "O", "P", "R", "S", "T", "U", "F", "KH", "TS",
      "CH", "SH", "SHCH", null, "YU", "YA",
      "a", "b", "v", "h","g", "d", "e","ie","zh", "z", "i","i", "i",
      "k", "l", "m", "n", "o", "p", "r", "s", "t", "u", "f", "kh", "ts",
      "ch", "sh", "shch", null, "iu", "ia"};

            for (int i = 0; i < ukr.Length; i++)
            {
                input = input.Replace(ukr[i], eng[i]);
            }

            return input;
        }


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


        public class ApiResponse
        {
            public ResponseData responseData { get; set; }
        }

        public class ResponseData
        {
            public string translatedText { get; set; }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                label3.Visible = false;

                long phoneNumber = long.Parse(textBox1.Text);
                textBox1.Text = phoneNumber.ToString();
                string lastName = textBox5.Text;
                string name = textBox6.Text;
                string surName = textBox7.Text;
                string gmail = "";
                string translatedLastName = Transliterate(lastName);
                string translatedName = Transliterate(name);
                translatedName = translatedName.ToLower();
                translatedLastName = translatedLastName.ToLower();

                gmail += $"{translatedLastName}.";
                gmail += $"{translatedName[0]}_{comboBox1.SelectedValue}23";
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
                    worksheet.Cells[lastRow + 1, 4].Value = phoneNumber.ToString();
                    worksheet.Cells[lastRow + 1, 5].Value = gmail;
                    worksheet.Cells[lastRow + 1, 6].Value = textBox9.Text;
                    package.Save();
                }
            }
            catch (FormatException)
            {
                label3.Visible = true;

            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string folderPath2 = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);

            // Формуємо повний шлях до файлу
            string filePath2 = Path.Combine(folderPath2, "data.xlsx");

            // Перевіряємо, чи існує файл, та якщо ні - створюємо новий
            if (!File.Exists(filePath2))
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
                    package.SaveAs(new FileInfo(filePath2));
                }
            }
            // Отримуємо шлях до файлу
            string folderPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            string filePath = Path.Combine(folderPath, "data.xlsx");

            // Відкриваємо папку з файлом Excel
            Process.Start("explorer.exe", $"/select,\"{filePath}\"");
        }

        
        private void Form1_Load(object sender, EventArgs e)
        {
            MessageBox.Show("Перед тим як створювати пошту,закрийте excel-файл!","Увага!", MessageBoxButtons.OK,MessageBoxIcon.Warning);
            label3.Visible = false;
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

        private void button2_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Перед тим як створювати пошту,закрийте excel-файл!", "Увага!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            try
            {
                string folderPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                // Формуємо повний шлях до файлу
                string filePath = Path.Combine(folderPath, "data.xlsx");

                // Перевіряємо, чи існує файл, та якщо ні - виводимо помилку
                if (!File.Exists(filePath))
                {
                    MessageBox.Show("Файл не знайдено!\nСтворіть першу пошту, після чого створиться автоматично новий excel-file","Увага!",MessageBoxButtons.OK,MessageBoxIcon.Information);
                    return;
                }

                // Відкриваємо файл з даними
                Process.Start(filePath);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Сталася помилка: {ex.Message}");
            }

        }
    }
}
