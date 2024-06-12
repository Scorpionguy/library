using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop;
using Word = Microsoft.Office.Interop.Word;
using Npgsql;
using System.Collections;

namespace library
{
    public partial class BooksDialog : Form
    {
        string connectionString = "server=localhost;username=postgres;database=library;port=5432;password=1234";
        string query = "select b.name, b.author, b.date, b.category, b.amount, p.name, p.country from books b left join publishing p on b.fk_publishing_id = p.publishing_id";
        NpgsqlConnection connection;
        public BooksDialog()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Создаем объект приложения Word
            Word.Application wordApp = new Word.Application();
            

            // Создаем новый документ
            Word.Document document = wordApp.Documents.Add();

            

            //// Освобождаем COM объекты
            //Marshal.ReleaseComObject(document);
            //Marshal.ReleaseComObject(wordApp);

            using (connection = new NpgsqlConnection(connectionString))
            {
                connection.Open();
                using (NpgsqlCommand command = new NpgsqlCommand(query, connection))
                {
                    List<string[]> data = new List<string[]>();
                    using (NpgsqlDataReader reader = command.ExecuteReader())
                    {
                        // Получаем количество столбцов и строк
                        int columnsCount = reader.FieldCount;
                        int rowsCount = 0;

                        while (reader.Read())
                        {
                            string[] row = new string[columnsCount];
                            for (int i = 0; i < columnsCount; i++)
                            {
                                row[i] = reader[i].ToString();
                            }
                            data.Add(row);
                        }

                        // Добавляем таблицу в документ Word (с учетом заголовков столбцов)
                        Word.Table table = document.Tables.Add(document.Range(0, 0), data.Count + 1, columnsCount);
                        table.Borders.Enable = 1; // Включаем границы таблицы

                        // Заполняем заголовки столбцов
                        for (int columnIndex = 0; columnIndex < columnsCount; columnIndex++)
                        {
                            table.Cell(1, columnIndex + 1).Range.Text = reader.GetName(columnIndex);
                        }

                        // Заполняем данные из списка
                        for (int rowIndex = 0; rowIndex < data.Count; rowIndex++)
                        {
                            for (int columnIndex = 0; columnIndex < columnsCount; columnIndex++)
                            {
                                table.Cell(rowIndex + 2, columnIndex + 1).Range.Text = data[rowIndex][columnIndex];
                            }
                        }
                    }
                }
                // Сохраняем документ (по желанию)
                string filePath = @"C:\Users\Fedor\OneDrive\Рабочий стол\Папка\ASD.docx";
                document.SaveAs2(filePath);

                // Закрываем документ и приложение Word
                document.Close();
                wordApp.Quit();
            }
            

            

        }
    }
}
