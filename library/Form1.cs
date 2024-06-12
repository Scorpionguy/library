using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using Npgsql;
using System.Collections;
using Microsoft.Office.Interop.Word;
using System.Runtime.InteropServices;


namespace library
{
    public partial class Form1 : Form
    {
        string connectionString = "server=localhost;username=postgres;database=library;port=5432;password=1234";
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            
            string query = "select b.name, b.author, b.date, b.category, b.amount, p.name, p.country from books b left join publishing p on b.fk_publishing_id = p.publishing_id";
            NpgsqlConnection connection;
            try
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
                }
                // Сохраняем документ (по желанию)

                using (SaveFileDialog saveFileDialog = new SaveFileDialog())
                {
                    saveFileDialog.Filter = "Word Document|*.docx";
                    saveFileDialog.Title = "Save Word Document";
                    saveFileDialog.DefaultExt = "docx";
                    saveFileDialog.AddExtension = true;

                    if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        // Сохраняем документ в выбранном пути
                        document.SaveAs2(saveFileDialog.FileName);
                    }
                }

                // Закрываем документ и приложение Word
                document.Close();
                wordApp.Quit();
            }

            // Исключение для ошибок
            catch (Exception ex)
            {
                MessageBox.Show($"Пожалуйста, повторите попытку! {ex.Message}", "Ошибка!");
            }
            
        }

            private void button2_Click(object sender, EventArgs e)
            {
            string query = "SELECT c.first_name as Имя, c.last_name as Фамилия,\r\nCOUNT(b.fk_client_id) AS Количество_Взятых_Книг\r\nFROM clients c\r\nJOIN bids b ON c.client_id = b.fk_client_id\r\nGROUP BY c.client_id, c.first_name;";
            NpgsqlConnection connection;
            try
            {
                // Создаем объект приложения Word
                Word.Application wordApp = new Word.Application();
                // Создаем новый документ
                Word.Document document = wordApp.Documents.Add();

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
                }

                // Сохраняем документ
                using (SaveFileDialog saveFileDialog = new SaveFileDialog())
                {
                    saveFileDialog.Filter = "Word Document|*.docx";
                    saveFileDialog.Title = "Save Word Document";
                    saveFileDialog.DefaultExt = "docx";
                    saveFileDialog.AddExtension = true;

                    if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        // Сохраняем документ в выбранном пути
                        document.SaveAs2(saveFileDialog.FileName);
                    }
                }

                // Закрываем документ и приложение Word
                document.Close();
                wordApp.Quit();
            }
            // Исключение для ошибок
            catch (Exception ex)
            {
                MessageBox.Show($"Что-то пошло не так, пожалуйста, повторите попытку! {ex.Message}", "Ошибка!");
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string query = "SELECT e.first_name as Имя, e.last_name as Фамилия,\r\nCOUNT(b.fk_inn) AS Количество_Взятых_Книг\r\nFROM employees e\r\nJOIN bids b ON e.inn = b.fk_inn\r\nGROUP BY e.inn, e.first_name;";
            NpgsqlConnection connection;
            try
            {


                // Создаем объект приложения Word
                Word.Application wordApp = new Word.Application();


                // Создаем новый документ
                Word.Document document = wordApp.Documents.Add();

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
                }
                // Сохраняем документ

                using (SaveFileDialog saveFileDialog = new SaveFileDialog())
                {
                    saveFileDialog.Filter = "Word Document|*.docx";
                    saveFileDialog.Title = "Save Word Document";
                    saveFileDialog.DefaultExt = "docx";
                    saveFileDialog.AddExtension = true;

                    if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        // Сохраняем документ в выбранном пути
                        document.SaveAs2(saveFileDialog.FileName);
                    }
                }

                // Закрываем документ и приложение Word
                document.Close();
                wordApp.Quit();
            }
            // Исключение для ошибок
            catch (Exception ex)
            {
                MessageBox.Show($"Что-то пошло не так, пожалуйста, повторите попытку! {ex.Message}", "Ошибка!");
            }
        }
    } 
}

