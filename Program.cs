using System;
using System.Windows.Forms;
using ExcelDataReader;
using System.Data;
using System.IO;
using System.Data.SqlClient;

namespace FotoGetConsole
{

    public static class Utils
    {
        public static DataTable GetPage(this DataSet dataSet, int pageNumber)
        {
            return dataSet.Tables[pageNumber - 1];
        }

        public static DataRow GetRow(this DataTable dataTable, int rowNumber)
        {
            return dataTable.Rows[rowNumber - 1];
        }

        public static string GetColumn(this DataRow row, char columnLetter)
        {
            int index = (int)columnLetter - (int)'A';
            var item = row[index];
            return item.ToString();
        }
    }

    internal class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            Console.WriteLine("Привет!\nНужно вытащить фото?\nДля работы потрубется файл Exel с FotoID. FotoID должны быть в столбце А на Листе1 без заголовка.\n\n1 - да, начать работу\n2 - нет, выход\n\nВведите номер команды: ");
            
            switch  (Convert.ToInt32(Console.ReadLine()))
            {
                case 1:

                    Console.WriteLine("\n\nНажмите любую кнопку, что бы выбрать файл Exel.");
                    Console.ReadKey(true);

                    OpenFileDialog ExcelFileName = new OpenFileDialog();
                    ExcelFileName.DefaultExt = ".xlsx";
                    ExcelFileName.Filter = "Excel|*.xlsx";
                    if (ExcelFileName.ShowDialog() == DialogResult.OK)
                    {
                        Console.WriteLine("Обрабатываемый файл: "+ExcelFileName.FileName);
                    }


                    Console.WriteLine("\n\nНажмите любую кнопку, что бы выбрать место сохранения фото.");
                    Console.ReadKey(true);
                    
                    FolderBrowserDialog SavePlase = new FolderBrowserDialog();
                    if (SavePlase.ShowDialog() == DialogResult.OK)
                    {
                        Console.WriteLine("Файлы сохранятся тут: " + SavePlase.SelectedPath);
                    }


                    string SQLInput = @"Data Source=WIN-VLHPBSGD7PP\SQLEXPRESS01;Initial Catalog=Test_01;Integrated Security=true";

                    using (StreamReader stream = new StreamReader(ExcelFileName.FileName))
                    {
                        using (var reader = ExcelReaderFactory.CreateReader(stream.BaseStream))
                        {
                            DataSet excellDataSet = reader.AsDataSet(new ExcelDataSetConfiguration());

                            int page = 1;
                            int rowsStart = 1;
                            char columnLetter = 'A';

                            int rowsCount = excellDataSet.GetPage(page).Rows.Count;

                            using (SqlConnection connection = new SqlConnection(SQLInput))
                            {
                                connection.Open();

                                for (int row = rowsStart; row <= rowsCount; row++)
                                {
                                    var value = excellDataSet.GetPage(page).GetRow(row).GetColumn(columnLetter);

                                    string SqlDostal = $"select Foto from PersonFoto1 where FotoID={value}";
                                    SqlCommand command = new SqlCommand(SqlDostal, connection);
                                    byte[] BB = (byte[])command.ExecuteScalar();
                                    File.WriteAllBytes(SavePlase.SelectedPath+$"\\{value}.jpg", BB);

                                }

                                connection.Close();

                            }

                        }
                    }

                    Console.WriteLine($"\n\nРабота закончена.\nНажмите любую кнопку для выхода.");
                    Console.ReadKey(true);
                    break;

                case 2:
                    break;
                default:
                    Console.WriteLine("Я не знаю такой команды.\nПожалуйста повторите запуск, если необходимо.");
                    Console.ReadKey(true);
                    break;

            }

            Console.Clear();

        }
    }
}
