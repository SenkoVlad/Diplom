using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

namespace allRows_Any_thing
{
    class Program
    {
        static void Main(string[] args)
        {
            List<string> allRows = new List<string>();
            List<string[]> allCountInRows = new List<string[]>();
            List<float> numbers_table_1 = new List<float>();
            List<float> numbers_table_2 = new List<float>();
            List<float> numbers_table_3 = new List<float>();
            List<float> numbers_table_4 = new List<float>();

            float number = 0;
            string pathToTemplatWeFile = @"D:\Diplom_\Diplom_08.05\Документы\Оценка НР 2018 210 (тест).doc";
            Word.Application wordApp = new Word.Application();
            wordApp.Visible = false;

            Word.Document document = wordApp.Documents.OpenNoRepairDialog(pathToTemplatWeFile);
            document.Activate();

            Word.Table table1 = document.Tables[1];
            Word.Table table2 = document.Tables[2];
            Word.Table table3 = document.Tables[3];
            Word.Table table4 = document.Tables[4];

            for (int i = 1; i <= table1.Rows.Count - 1; i++)
            {
                allRows.Add(table1.Cell(i, table1.Rows[i].Cells.Count).Range.Text.Replace("\r", " ").Replace("\a", " ").Replace("\n", " "));
            }

            for (int i = 0; i < allRows.Count; i++)
            {
                allCountInRows.Add(allRows[i].Split(' '));
            }

            for (int i = 0; i < allCountInRows.Count; i++)
            {
                if (i == 1)
                {
                    for (int j = 0; j < allCountInRows[i].Length; j++)
                    {
                        if (float.TryParse(allCountInRows[i][j], out number))
                        {
                            numbers_table_1.Add(number);
                        }
                    }
                }
                else
                {
                    if (float.TryParse(allCountInRows[i][0], out number))
                    {
                        numbers_table_1.Add(number);
                    }
                }

            }
            allCountInRows.Clear();
            allRows.Clear();

            for (int i = 1; i <= table2.Rows.Count - 1; i++)
            {
                try
                {
                    allRows.Add(table2.Cell(i, 5).Range.Text.Replace("\r", " ").Replace("\a", " ").Replace("\n", " "));
                }
                catch (Exception ex)
                {
                    allRows.Add(table2.Cell(i, 4).Range.Text.Replace("\r", " ").Replace("\a", " ").Replace("\n", " "));
                }
            }

            for (int i = 0; i < allRows.Count; i++)
            {
                allCountInRows.Add(allRows[i].Split(' '));
            }

            for (int i = 0; i < allCountInRows.Count; i++)
            {
                if (float.TryParse(allCountInRows[i][0], out number))
                {
                    numbers_table_2.Add(number);
                }
            }
            allCountInRows.Clear();
            allRows.Clear();


            for (int i = 1; i <= table3.Rows.Count - 1; i++)
            {
                try
                {
                    allRows.Add(table3.Cell(i, 5).Range.Text.Replace("\r", " ").Replace("\a", " ").Replace("\n", " "));
                }
                catch (Exception ex)
                {
                    allRows.Add(table3.Cell(i, 4).Range.Text.Replace("\r", " ").Replace("\a", " ").Replace("\n", " "));
                }
            }

            for (int i = 0; i < allRows.Count; i++)
            {
                allCountInRows.Add(allRows[i].Split(' '));
            }

            for (int i = 0; i < allCountInRows.Count; i++)
            {
                if (float.TryParse(allCountInRows[i][0], out number))
                {
                    numbers_table_3.Add(number);
                }
            }
            allCountInRows.Clear();
            allRows.Clear();


            for (int i = 1; i <= table4.Rows.Count - 1; i++)
            {
                try
                {
                    allRows.Add(table4.Cell(i, 5).Range.Text.Replace("\r", " ").Replace("\a", " ").Replace("\n", " "));
                }
                catch (Exception ex)
                {
                    allRows.Add(table4.Cell(i, 4).Range.Text.Replace("\r", " ").Replace("\a", " ").Replace("\n", " "));
                }
            }

            for (int i = 0; i < allRows.Count; i++)
            {
                allCountInRows.Add(allRows[i].Split(' '));
            }

            for (int i = 0; i < allCountInRows.Count; i++)
            {
                if (float.TryParse(allCountInRows[i][0], out number))
                {
                    numbers_table_4.Add(number);
                }
            }
            allCountInRows.Clear();
            allRows.Clear();

            Console.WriteLine(numbers_table_1.Sum());
            Console.WriteLine(numbers_table_2.Sum());
            Console.WriteLine(numbers_table_3.Sum());
            Console.WriteLine(numbers_table_4.Sum());

            document.Close();
            wordApp.Quit();
        }
    }
}
