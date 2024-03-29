using System;
using OfficeOpenXml;
using System.Windows.Forms;
using System.Management;
using System.IO;

namespace InventoryManagementApp
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // Встановлення контексту ліцензування

        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Отримуємо інформацію про комп'ютер
            string computerType = GetComputerType();
            string operatingSystem = GetOperatingSystem();
            string userName = Environment.UserName; // Отримуємо ім'я користувача
            string memory = GetMemory(); // Отримуємо інформацію про обсяг оперативної пам'яті
            string officeVersion = GetMicrosoftOfficeVersion(); // Отримуємо інформацію про версію офісного пакету
            string itEnterprise = CheckITEnterprise(); // Перевіряємо, чи встановлено IT Enterprise

            // Виводимо отриману інформацію в DataGridView
            dataGridView1.Rows.Add(
                Environment.MachineName, // Ім'я пристрою
                (computerType == "1") ? "Стаціонарний ПК" : "Ноутбук", // Тип ПК
                operatingSystem, // Назва операційної системи
                Environment.Is64BitOperatingSystem ? "64-біт" : "32-біт", // Разрядність
                userName, // Ім'я користувача
                memory, // Оперативна пам'ять
                officeVersion, // Версія офісного пакету
                itEnterprise // Наявність IT Enterprise
            );
            ExportToExcel();
            MessageBox.Show("Дані експортовано до файлу Excel.");
        }

        private string GetComputerType()
        {
            string query = "SELECT PCSystemType FROM Win32_ComputerSystem";
            ManagementObjectSearcher searcher = new ManagementObjectSearcher(query);
            ManagementObjectCollection collection = searcher.Get();
            foreach (ManagementObject obj in collection)
            {
                return obj["PCSystemType"].ToString();
            }
            return "Невідомо";
        }

        private string GetOperatingSystem()
        {
            string query = "SELECT Caption, OSArchitecture FROM Win32_OperatingSystem";
            ManagementObjectSearcher searcher = new ManagementObjectSearcher(query);
            ManagementObjectCollection collection = searcher.Get();
            foreach (ManagementObject obj in collection)
            {
                string caption = obj["Caption"].ToString();
                string architecture = obj["OSArchitecture"].ToString();
                return $"{caption} ({architecture}-біт)";
            }
            return "Невідомо";
        }

        private string GetMemory()
        {
            string query = "SELECT TotalPhysicalMemory FROM Win32_ComputerSystem";
            ManagementObjectSearcher searcher = new ManagementObjectSearcher(query);
            ManagementObjectCollection collection = searcher.Get();
            foreach (ManagementObject obj in collection)
            {
                ulong memorySize = Convert.ToUInt64(obj["TotalPhysicalMemory"]);
                double memorySizeInGB = memorySize / (1024.0 * 1024 * 1024); // Переводимо байти в гігабайти
                return $"{memorySizeInGB:F2} ГБ";
            }
            return "Невідомо";
        }

        private string GetMicrosoftOfficeVersion()
        {
            try
            {
                // Перевіряємо, чи встановлений Microsoft 365
                string microsoft365Version = GetMicrosoft365Version();
                if (!string.IsNullOrEmpty(microsoft365Version))
                {
                    return "Microsoft 365";
                }

                // Перевіряємо, чи встановлений Microsoft Office
                string officeVersion = GetOfficeVersion();
                if (!string.IsNullOrEmpty(officeVersion))
                {
                    return "Microsoft Office";
                }

                // Перевіряємо, чи встановлений LibreOffice
                string libreOfficeVersion = GetLibreOfficeVersion();
                if (!string.IsNullOrEmpty(libreOfficeVersion))
                {
                    return "LibreOffice";
                }

                // Перевіряємо, чи встановлений Office 365
                string office365Version = GetOffice365Version();
                if (!string.IsNullOrEmpty(office365Version))
                {
                    return "Office 365";
                }
            }
            catch (Exception ex)
            {
                // Обробка помилок доступу до реєстру або відсутності ключа
                return "Невідомо";
            }

            // Якщо назва продукту не визначена, повертаємо "Невідомо"
            return "Невідомо";
        }

        private string GetMicrosoft365Version()
        {
            // Шлях до виконуваного файлу Microsoft 365
            string path = @"C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE"; // Приклад для Outlook, можна змінити для інших програм Microsoft 365

            // Отримуємо версію за допомогою властивості VersionInfo
            System.Diagnostics.FileVersionInfo info = System.Diagnostics.FileVersionInfo.GetVersionInfo(path);
            return info.FileVersion;
        }

        private string GetOfficeVersion()
        {
            // Шлях до виконуваного файлу Office
            string path = @"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE"; // Приклад для Word, можна змінити для інших програм Office

            // Отримуємо версію за допомогою властивості VersionInfo
            System.Diagnostics.FileVersionInfo info = System.Diagnostics.FileVersionInfo.GetVersionInfo(path);
            return info.FileVersion;
        }

        private string GetLibreOfficeVersion()
        {
            // Шлях до виконуваного файлу LibreOffice
            string path = @"C:\Program Files\LibreOffice\program\soffice.bin"; // Приклад для LibreOffice Writer, можна змінити для інших програм LibreOffice

            // Отримуємо версію за допомогою властивості VersionInfo
            System.Diagnostics.FileVersionInfo info = System.Diagnostics.FileVersionInfo.GetVersionInfo(path);
            return info.FileVersion;
        }

        private string GetOffice365Version()
        {
            // Шлях до виконуваного файлу Office 365
            string path = @"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE"; // Приклад для Word, можна змінити для інших програм Office 365

            // Отримуємо версію за допомогою властивості VersionInfo
            System.Diagnostics.FileVersionInfo info = System.Diagnostics.FileVersionInfo.GetVersionInfo(path);
            return info.FileVersion;
        }
        // Перевіряємо, чи встановлено IT Enterprise
        private string CheckITEnterprise()
        {
            string directoryPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            directoryPath = Path.Combine(directoryPath, @"Microsoft\Windows\Start Menu\Programs\IT");

            if (Directory.Exists(directoryPath))
            {
                return "Клієнт IT-Enterprise встановлений";
            }
            else
            {
                return "Клієнт IT-Enterprise не встановлений";
            }
        }
        private void ExportToExcel()
        {
            // Створення нового файлу Excel
            FileInfo file = new FileInfo("InventoryData.xlsx");
            using (ExcelPackage package = new ExcelPackage(file))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Inventory");

                // Заповнення заголовків
                for (int i = 0; i < dataGridView1.Columns.Count; i++)
                {
                    worksheet.Cells[1, i + 1].Value = dataGridView1.Columns[i].HeaderText;
                }

                // Заповнення даними
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    for (int j = 0; j < dataGridView1.Columns.Count; j++)
                    {
                        worksheet.Cells[i + 2, j + 1].Value = dataGridView1.Rows[i].Cells[j].Value;
                    }
                }

                // Збереження файлу
                package.Save();
            }
        }
    }
}
