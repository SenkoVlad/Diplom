using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using Word = Microsoft.Office.Interop.Word;
using System.Threading;

namespace Science_0._01
{
    public partial class Form_reports : Form
    {
        //Для поединения с БД
        private OleDbConnection connection = new OleDbConnection();

        //Для выполненения команды
        private OleDbCommand command = new OleDbCommand();
        private OleDbCommand command2 = new OleDbCommand();
        private OleDbCommand command3 = new OleDbCommand();

        //Для агрегации встравляевых строк
        private List<string> insertString = new List<string>();
        //Вставляемая строка
        private string insertText = "";

        //Для чтения данных из таблицы
        private OleDbDataReader reader = null;
        private OleDbDataReader reader2 = null;
        private OleDbDataReader reader3 = null;

        //Подсчёт строк результата
        private int countRow = 0;

        //Словарь для агрегации строк запроса
        private Dictionary<string, string> distinctColumn = new Dictionary<string, string>();

        //Для вставки в word файл
        private string pathToTemplateFile =  Application.StartupPath.ToString() + @"\Оценка НР 2018 210 (Шаблон).doc";
        private string pathToTemplFile = Application.StartupPath.ToString() + @"\1 квартал 210 (Шаблон).doc";

        //Количество ППС
        private string PPS_count = "";

        //Для кнопки выбрать все!
        private bool all_cheaked = false;
        private bool all_cheaked_2 = false;


        //Для вставки в документ
        private string text_table_3_1_2 = "";

        private string text_table_3_2_1 = "";
        private string text_table_3_2_2 = "";
        private string text_table_3_2_3_1 = "";
        private string text_table_3_2_3_2 = "";
        private string text_table_3_2_4_1 = "";
        private string text_table_3_2_4_2 = "";

        private string text_table_3_2_5 = "";

        private string text_table_3_3_1_1 = "";
        private string text_table_3_3_1_2 = "";
        private string text_table_3_3_1_3 = "";
        private string text_table_3_3_1_4 = "";
        private string text_table_3_3_1_5 = "";
        private string text_table_3_3_1_6 = "";

        private string text_table_3_3_2_1 = "";
        private string text_table_3_3_2_2 = "";
        private string text_table_3_3_2_3 = "";
        private string text_table_3_3_2_4 = "";
        private string text_table_3_3_2_5 = "";
        private string text_table_3_3_2_6 = "";
        private string text_table_3_3_2_7 = "";
        private string text_table_3_3_2_8 = "";
        private string text_table_3_3_2_9 = "";
        private string text_table_3_3_2_10 = "";
        private string text_table_3_3_2_11 = "";
        private string text_table_3_3_2_12 = "";

        private string text_table_3_3_3_1 = "";
        private string text_table_3_3_3_2 = "";
        private string text_table_3_3_3_3 = "";
        private string text_table_3_3_3_4 = "";

        private string text_table_3_4_1_1 = "";
        private string text_table_3_4_1_2 = "";
        private string text_table_3_4_1_3 = "";
        private string text_table_3_4_1_4 = "";
        private string text_table_3_4_1_5 = "";

        private string text_table_3_4_3_1 = "";
        private string text_table_3_4_3_2 = "";
        private string text_table_3_4_3_3 = "";
        private string text_table_3_4_3_4 = "";
        private string text_table_3_4_3_5 = "";
        private string text_table_3_4_3_6 = "";
        private string text_table_3_4_3_7 = "";
        private string text_table_3_4_3_8 = "";

        private string text_table_3_4_4_1 = "";
        private string text_table_3_4_4_2 = "";
        private string text_table_3_4_4_3 = "";
        private string text_table_3_4_4_4 = "";
        private string text_table_3_4_4_5 = "";
        private string text_table_3_4_4_6 = "";
        private string text_table_3_4_4_7 = "";


        private int score_table_3_1_2 = 0;

        private int score_table_3_2_4_2 = 0;

        private int score_table_3_2_2 = 0;
        private int score_table_3_2_3_1 = 0;
        private int score_table_3_2_3_2 = 0;
        private int score_table_3_2_4_1 = 0;
        private int score_table_3_2_5 = 0;

        private int score_table_3_3_1_1 = 0;
        private int score_table_3_3_1_2 = 0;
        private int score_table_3_3_1_3 = 0;
        private int score_table_3_3_1_4 = 0;
        private int score_table_3_3_1_5 = 0;
        private int score_table_3_3_1_6 = 0;

        private int score_table_3_3_3_1 = 0;
        private int score_table_3_3_3_2 = 0;
        private int score_table_3_3_3_3 = 0;
        private int score_table_3_3_3_4 = 0;

        private int score_table_3_4_1_1 = 0;
        private int score_table_3_4_1_2 = 0;
        private int score_table_3_4_1_3 = 0;
        private int score_table_3_4_1_4 = 0;
        private int score_table_3_4_1_5 = 0;

        private int score_table_3_4_3_5 = 0;
        private int score_table_3_4_3_6 = 0;
        private int score_table_3_4_3_7 = 0;
        private int score_table_3_4_3_8 = 0;

        private int score_table_3_4_4_1 = 0;
        private int score_table_3_4_4_2 = 0;
        private int score_table_3_4_4_3 = 0;
        private int score_table_3_4_4_4 = 0;



        public Form_reports()
        {
            InitializeComponent();
            connection.ConnectionString = Form_login.connectString;

            command.Connection = connection;
            command2.Connection = connection;
            command3.Connection = connection;

            PPS_count = calcNIRScore("id_person", "(select distinct id_person from person)");
        }

        //Семинары и конференции на кафедре
        private void ConfAndSemPPS(ref string text_table_3_1_2, ref int score_table_3_1_2)
        {
            distinctColumn = calcOfIndicatorsConf("name_conference", "seminar", "status='Рес  публиканский' and place_conference='кафедра 210'", 10);
            if (distinctColumn["text"] != "")
            {
                text_table_3_1_2 = distinctColumn["text"] + Environment.NewLine;
                score_table_3_1_2 += int.Parse(distinctColumn["count"]);
            }

            distinctColumn = calcOfIndicatorsConf("name_conference", "seminar", "status='Вузовский' and place_conference='кафедра 210'", 5);
            if (distinctColumn["text"] != "")
            {
                text_table_3_1_2 += distinctColumn["text"] + Environment.NewLine;
                score_table_3_1_2 += int.Parse(distinctColumn["count"]);
            }


            distinctColumn = calcOfIndicatorsConf("name_conference", "conference", "status='Международный' and place_conference='кафедра 210'", 30);
            if (distinctColumn["text"] != "")
            {
                text_table_3_1_2 += distinctColumn["text"] + Environment.NewLine;
                score_table_3_1_2 += int.Parse(distinctColumn["count"]);
            }

            distinctColumn = calcOfIndicatorsConf("name_conference", "conference", "status='Республиканский' and place_conference='кафедра 210'", 20);
            if (distinctColumn["text"] != "")
            {
                text_table_3_1_2 += distinctColumn["text"] + Environment.NewLine;
                score_table_3_1_2 += int.Parse(distinctColumn["count"]);
            }

            distinctColumn = calcOfIndicatorsConf("name_conference", "conference", "status='Вузовский' and place_conference='кафедра 210'", 10);
            if (distinctColumn["text"] != "")
            {
                text_table_3_1_2 += distinctColumn["text"] + Environment.NewLine;
                score_table_3_1_2 += int.Parse(distinctColumn["count"]);
            }

            text_table_3_1_2 = score_table_3_1_2.ToString() + Environment.NewLine + Environment.NewLine + text_table_3_1_2;
        }

        //Поиск соискателей
        private Dictionary<string, string> ScienceLeaderCount(string conditionColumn)
        {
            try
            {
                connection.Open();
                Dictionary<string, string> distinctColumn = new Dictionary<string, string>();

                insertText = "";
                int count = 0;
                command.CommandText = "select distinct leader,led_by from scientific_leaders where led_by='" + conditionColumn + "';";
                reader = command.ExecuteReader();

                while (reader.Read())
                {
                    insertText += reader[0].ToString() + Environment.NewLine;
                    count++;
                }
                distinctColumn.Add("text", insertText);
                distinctColumn.Add("count", count.ToString());

                connection.Close();
                reader.Close();

                return distinctColumn;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                connection.Close();
                connection.Dispose();
                return null;
            }
        }

        //Степень участия ППС в разработке учебников, учебных пособий, моногарфий
        private Dictionary<string, string> BooksCalc(string conditionColumn)
        {
            try
            {
                connection.Open();
                Dictionary<string, string> distinctColumn = new Dictionary<string, string>();

                insertText = "";

                command.CommandText = "select title_textbook,title_edition from (select  distinct title_textbook,title_edition from textbooks where book_view='" + conditionColumn + "');";
                reader = command.ExecuteReader();

                while (reader.Read())
                {
                    command2.CommandText = "select person.Fam,person.Name,person.mid_name from textbooks inner join person on textbooks.id_person = person.id_person where book_view='" + conditionColumn + "' and title_textbook='" + reader[0].ToString() + "' and title_edition='" + reader[1].ToString() + "';";

                    reader2 = command2.ExecuteReader();
                    if (!distinctColumn.ContainsKey("(" + reader[0].ToString() + ", " + reader[1].ToString() + ")"))
                        distinctColumn.Add("(" + reader[0].ToString() + ", " + reader[1].ToString() + ")", "");

                    while (reader2.Read())
                    {
                        insertText += reader2[0].ToString() + " " + reader2[1].ToString()[0] + "." + reader2[2].ToString()[0] + "., ";
                    }

                    distinctColumn["(" + reader[0].ToString() + ", " + reader[1].ToString() + ")"] += insertText.Substring(0, insertText.Length - 2);
                    insertText = "";
                    if (!reader2.IsClosed)
                        reader2.Close();
                }
                connection.Close();
                reader2.Close();
                reader.Close();

                distinctColumn.Add("count", calcNIRScore("sum(number_sheets)", "(select distinct title_textbook, title_edition, number_sheets from textbooks where book_view='" + conditionColumn + "')"));

                return distinctColumn;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                connection.Close();
                connection.Dispose();
                return null;
            }
        }


        //Все конференции и семинары
        private Dictionary<string, string> calcOfIndicatorsConf(string columnsForRead, string tableNane, string condition, int scoreFactor)
        {
            try
            {
                Dictionary<string, string> distinctColumn = new Dictionary<string, string>();

                connection.Open();

                countRow = 0;
                insertText = "";
                command.CommandText = "select distinct " + columnsForRead + " from " + tableNane + " inner join conference_list on conference_list.id_record_conf = " + tableNane + ".id_record_conf where " + condition + ";";
                reader = command.ExecuteReader();
                while (reader.Read())
                {
                    countRow = 1;
                    insertText += reader[0].ToString() + Environment.NewLine;
                }
                distinctColumn.Add("text", insertText);
                distinctColumn.Add("count", (countRow * scoreFactor).ToString());

                reader.Close();
                connection.Close();

                return distinctColumn;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                connection.Close();
                connection.Dispose();
                return null;
            }

        }

        //Результативность выполненных НИР 
        private Dictionary<string, string> calcAllNir(string columnsForRead, string tableNane, string condition)
        {
            try
            {
                connection.Open();
                insertText = "";
                Dictionary<string, string> distinctColumn = new Dictionary<string, string>();

                command.CommandText = "select " + columnsForRead + " from " + tableNane + " where " + condition + ";";
                reader = command.ExecuteReader();
                while (reader.Read())
                {
                    insertText += '"' + reader[0].ToString() + "\", ";
                }
                reader.Close();
                connection.Close();

                distinctColumn.Add("text", insertText.Substring(0, insertText.Length - 2));
                distinctColumn.Add("count", calcNIRScore("count(*)", "(select distinct  title_reasearch from reasearch_work where " + condition + ")"));

                return distinctColumn;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                connection.Close();
                connection.Dispose();
                return null;
            }

        }

        //Количество ППС делавшие НИР
        private string calcNIRScore(string columnsForRead, string tableName)
        {
            try
            {
                connection.Open();
                string countPerson = "";
                insertText = "";

                command.CommandText = "select " + columnsForRead + " from " + tableName + ";";
                reader = command.ExecuteReader();
                while (reader.Read())
                {
                    countPerson = reader[0].ToString();
                }
                reader.Close();
                connection.Close();

                insertText = countPerson;
                return insertText;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                connection.Close();
                connection.Dispose();
                return null;
            }

        }

        //Степень участия научных работников высшей квалификации из числа ППС в подготовке научных работников высшей квалификации 
        private Dictionary<string, string> calcAllReviews(string columnsForRead, string tableNane, string condition)
        {
            try
            {
                connection.Open();
                insertText = "";
                int countRows = 0;
                Dictionary<string, string> distinctColumn = new Dictionary<string, string>();

                command.CommandText = "select " + columnsForRead + " from " + tableNane + " where " + condition + ";";
                reader = command.ExecuteReader();
                while (reader.Read())
                {
                    countRows++;
                    insertText += reader[1].ToString().Split(' ')[0] + " " + reader[1].ToString().Split(' ')[1][0] + "." + " " + reader[1].ToString().Split(' ')[2][0] + "." + Environment.NewLine + reader[0].ToString() + Environment.NewLine;
                }
                reader.Close();
                connection.Close();

                distinctColumn.Add("text", insertText);
                distinctColumn.Add("count", countRows.ToString());

                return distinctColumn;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                connection.Close();
                connection.Dispose();
                return null;
            }

        }



        //Степень участия ППС в изобретательской и рационализаторской работе
        private Dictionary<string, string> calcInventions(string condition)
        {
            try
            {
                connection.Open();
                Dictionary<string, string> distinctColumn = new Dictionary<string, string>();


                insertText = "";

                command.CommandText = "select title_invention from (select  distinct title_invention from inventions where " + condition + ");";
                reader = command.ExecuteReader();

                while (reader.Read())
                {
                    command2.CommandText = "select person.Fam,person.Name,person.mid_name from inventions inner join person on inventions.id_person = person.id_person where " + condition + " and title_invention='" + reader[0].ToString() + "';";

                    reader2 = command2.ExecuteReader();
                    distinctColumn.Add(reader[0].ToString(), "");

                    while (reader2.Read())
                    {
                        insertText += reader2[0].ToString() + " " + reader2[1].ToString()[0] + "." + reader2[2].ToString()[0] + "., ";
                    }

                    distinctColumn[reader[0].ToString()] += insertText.Substring(0, insertText.Length - 2);
                    insertText = "";
                    if (!reader2.IsClosed)
                        reader2.Close();
                }
                reader2.Close();
                reader.Close();
                connection.Close();

                distinctColumn.Add("count", calcNIRScore("count(*)", "(select distinct title_invention from inventions where " + condition + ")"));

                return distinctColumn;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                connection.Close();
                connection.Dispose();
                return null;
            }

        }

        //Степень участия ППС на выставках и конкурсах научных и творческих работ 
        private string calcPpsExhibitions(string columnsForRead, string tableName, string condition, int scoreFactor)
        {
            try
            {
                connection.Open();

                countRow = 0;
                insertText = "";

                command.CommandText = "select " + columnsForRead + " from " + tableName + " where " + condition + ";";
                reader = command.ExecuteReader();
                while (reader.Read())
                {
                    countRow++;
                    insertText += reader[0].ToString() + Environment.NewLine;
                }
                insertText = (Convert.ToInt32(countRow * scoreFactor / float.Parse(PPS_count))).ToString() + Environment.NewLine + Environment.NewLine + insertText;

                reader.Close();
                connection.Close();

                return insertText;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                connection.Close();
                connection.Dispose();
                return null;
            }

        }



        //Степень участия ППС в публикации результатов научных исследований (Статьи) 
        private Dictionary<string, string> CalcPpsArticles(string conditionColumn)
        {
            try
            {
                connection.Open();
                Dictionary<string, string> distinctColumn = new Dictionary<string, string>();


                countRow = 0;
                insertText = "";

                command.CommandText = "select title_article,articles_list.title_edition from (select  distinct title_article, articles_list.title_edition from articles inner join articles_list on articles_list.id_article_record=articles.id_article_record where level_article='" + conditionColumn + "');";
                reader = command.ExecuteReader();

                while (reader.Read())
                {
                    countRow++;
                    command2.CommandText = " select distinct id_person  from  articles  inner join articles_list  on articles_list.id_article_record=articles.id_article_record where level_article='" + conditionColumn + "' and title_article='" + reader[0].ToString() + "' and articles_list.title_edition='" + reader[1].ToString() + "';";
                    reader2 = command2.ExecuteReader();

                    if (!distinctColumn.ContainsKey("(" + reader[0].ToString() + ", " + reader[1].ToString() + ")"))
                        distinctColumn.Add("(" + reader[0].ToString() + ", " + reader[1].ToString() + ")", "");

                    while (reader2.Read())
                    {
                        command3.CommandText = " select Fam, Name, mid_name  from  person  where id_person=" + reader2[0].ToString() + ";";
                        reader3 = command3.ExecuteReader();
                        while (reader3.Read())
                        {
                            insertText += reader3[0].ToString() + " " + reader3[1].ToString()[0] + "." + reader3[2].ToString()[0] + "., ";
                        }
                        if (!reader3.IsClosed)
                            reader3.Close();
                    }

                    distinctColumn["(" + reader[0].ToString() + ", " + reader[1].ToString() + ")"] += insertText.Substring(0, insertText.Length - 2);
                    insertText = "";
                    if (!reader2.IsClosed)
                        reader2.Close();
                }
                reader2.Close();
                reader.Close();
                connection.Close();

                distinctColumn.Add("count", calcNIRScore("count(*)", "(select distinct id_person from articles where level_article='" + conditionColumn + "')"));

                return distinctColumn;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                connection.Close();
                connection.Dispose();
                return null;
            }
        }


        //Степень участия ППС в апробации результатов научных исследований на научных конференциях
        //(Конференции) 
        private Dictionary<string, string> calcPpsConf(string conditionColumn)
        {
            try
            {
                connection.Open();
                Dictionary<string, string> distinctColumn = new Dictionary<string, string>();


                countRow = 0;
                insertText = "";
                command.CommandText = "select distinct name_conference,conference_list.id_record_conf from conference inner join conference_list on conference_list.id_record_conf= conference.id_record_conf  where status='" + conditionColumn + "';";
                reader = command.ExecuteReader();

                while (reader.Read())
                {
                    countRow++;
                    command2.CommandText = "select distinct id_person from conference inner join conference_list on conference.id_record_conf = conference_list.id_record_conf where status='" + conditionColumn + "' and conference_list.id_record_conf=" + reader[1].ToString() + ";";

                    reader2 = command2.ExecuteReader();
                    if (!distinctColumn.ContainsKey("(" + reader[0].ToString() + ")"))
                        distinctColumn.Add("(" + reader[0].ToString() + ")", "");

                    while (reader2.Read())
                    {
                        command3.CommandText = " select Fam, Name, mid_name  from  person  where id_person=" + reader2[0].ToString() + ";";
                        reader3 = command3.ExecuteReader();
                        while (reader3.Read())
                        {
                            insertText += reader3[0].ToString() + " " + reader3[1].ToString()[0] + "." + reader3[2].ToString()[0] + "., ";
                        }
                        if (!reader3.IsClosed)
                            reader3.Close();
                    }
                    distinctColumn["(" + reader[0].ToString() + ")"] += insertText.Substring(0, insertText.Length - 2);

                    insertText = "";
                    if (!reader2.IsClosed)
                        reader2.Close();

                }
                connection.Close();
                reader2.Close();
                reader.Close();

                distinctColumn.Add("count", calcNIRScore("count(*)", "(select distinct id_person from conference where status='" + conditionColumn + "')"));

                return distinctColumn;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                connection.Close();
                connection.Dispose();
                return null;
            }
        }


        //Участие курсантов в конференциях
        private Dictionary<string, string> calcKursantConf(string conditionColumn)
        {
            try
            {
                connection.Open();
                Dictionary<string, string> distinctColumn = new Dictionary<string, string>();

                countRow = 0;

                command.CommandText = "select title_conf_kursant,FIO from conf_kursant where status='" + conditionColumn + "';";
                reader = command.ExecuteReader();

                while (reader.Read())
                {
                    countRow++;
                    if (!distinctColumn.ContainsKey("(" + reader[0].ToString() + ")"))
                    {
                        distinctColumn.Add("(" + reader[0].ToString() + ")", "");
                    }
                    distinctColumn["(" + reader[0].ToString() + ")"] += reader[1].ToString() + ", ";
                }

                connection.Close();
                reader.Close();

                distinctColumn.Add("count", calcNIRScore("count(*)", "(select distinct FIO from conf_kursant where status='" + conditionColumn + "')"));

                return distinctColumn;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                connection.Close();
                connection.Dispose();
                return null;
            }
        }

        //Подготовка научных работников высшей квалификации в докторантурах, адъюнктурах и в форме соискательства
        //укомплектованность должностей подлежащих замещению профессорами
        //укомплектованность должностей подлежащих замещению доцентами
        //удельный вес докторов наук
        //удельный вес кандидатов наук
        private Dictionary<string, string> CalcPpsWeight(string columnsForRead, string tableNane, string condition)
        {
            try
            {
                connection.Open();
                Dictionary<string, string> distinctColumn = new Dictionary<string, string>();
                insertText = "";
                int countRow = 0;

                command.CommandText = "select " + columnsForRead + " from " + tableNane + " where " + condition + ";";
                reader = command.ExecuteReader();
                while (reader.Read())
                {
                    countRow++;
                    insertText += reader[0].ToString() + " " + reader[1].ToString()[0] + "." + reader[2].ToString()[0] + "., ";
                }
                distinctColumn.Add("count", countRow.ToString());

                if (insertText.Length > 0)
                    distinctColumn.Add("FIO", insertText.Substring(0, insertText.Length - 2));
                else
                    distinctColumn.Add("FIO", "");

                reader.Close();
                connection.Close();

                return distinctColumn;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                connection.Close();
                connection.Dispose();
                return null;
            }
        }

        //Степень участия научных работников высшей квалификации из числа ППС в подготовке научных работников высшей квалификации 
        private Dictionary<string, string> CalcPpsExpertise(string columnsForRead, string tableNane, string condition)
        {
            try
            {
                connection.Open();
                Dictionary<string, string> distinctColumn = new Dictionary<string, string>();
                insertText = "";
                int countRow = 0;

                command.CommandText = "select " + columnsForRead + " from " + tableNane + " where " + condition + ";";
                reader = command.ExecuteReader();
                while (reader.Read())
                {
                    countRow++;
                    insertText += reader[0].ToString() + " " + reader[1].ToString()[0] + "." + reader[2].ToString()[0] + ",";
                }
                distinctColumn.Add("count", countRow.ToString());
                distinctColumn.Add("FIO", insertText.Substring(0, insertText.Length - 1));


                reader.Close();
                connection.Close();

                return distinctColumn;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                connection.Close();
                connection.Dispose();
                return null;
            }
        }

        //Степень участия ППС в работе по проведению научной экспертизы уставных документов, учебников и учебных пособий
        private Dictionary<string, string> CalcPpsReviews(string columnsForRead, string tableNane, string condition)
        {
            try
            {
                connection.Open();
                Dictionary<string, string> distinctColumn = new Dictionary<string, string>();
                insertText = "";
                int countRow = 0;

                command.CommandText = "select distinct title_review,person from reviews where tag='Документ/Учебник/Пособие';";
                reader = command.ExecuteReader();
                while (reader.Read())
                {
                    countRow++;
                    insertText += reader[1].ToString() + "\n" + reader[0].ToString() + "\n";
                }
                distinctColumn.Add("count", countRow.ToString());
                distinctColumn.Add("text", insertText);


                reader.Close();
                connection.Close();

                return distinctColumn;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                connection.Close();
                connection.Dispose();
                return null;
            }
        }

        //тезисов докладов в сборниках материалов научных конференций
        private Dictionary<string, string> CalcThesisConf()
        {
            try
            {
                connection.Open();
                Dictionary<string, string> distinctColumn = new Dictionary<string, string>();
                insertText = "";
                int countRow = 0;
                int allRow = 0;
                command.CommandText = "select name_conference,title_report  from conference inner join conference_list on conference_list.id_record_conf=conference.id_record_conf group by name_conference,title_report;";
                reader = command.ExecuteReader();
                while (reader.Read())
                {
                    allRow++;
                    countRow++;
                    if (!distinctColumn.ContainsKey(reader[0].ToString()))
                    {
                        countRow = 1;
                        distinctColumn.Add(reader[0].ToString(), countRow.ToString());
                    }
                    else
                    {
                        distinctColumn[reader[0].ToString()] = countRow.ToString();
                    }
                }
                distinctColumn.Add("count", allRow.ToString());


                reader.Close();
                connection.Close();

                return distinctColumn;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                connection.Close();
                connection.Dispose();
                return null;
            }
        }


        private void ReplaceWordStub(string stubToReplace, string text, Word.Document wordDocument)
        {
            var range = wordDocument.Content;
            range.Find.ClearFormatting();
            range.Find.Execute(FindText: stubToReplace, ReplaceWith: text);
        }

        private void Form_reports_FormClosed(object sender, FormClosedEventArgs e)
        {
            System.Environment.Exit(0); // Приложение завершается и возвращает ОС указанное параметром значение
        }

        private void ScienceLeader(ref string text_table_3_4_4_5, ref string text_table_3_4_4_6, ref string text_table_3_4_4_7)
        {
            distinctColumn = ScienceLeaderCount("дн");
            text_table_3_4_4_5 = Convert.ToInt32(float.Parse(distinctColumn["count"]) / float.Parse(calcNIRScore("count(*)", "(select id_person from person where academic_degree='дн')")) * 400).ToString()
                + Environment.NewLine + Environment.NewLine + distinctColumn["text"];

            distinctColumn = ScienceLeaderCount("кн");
            text_table_3_4_4_6 = Convert.ToInt32(float.Parse(distinctColumn["count"]) / float.Parse(calcNIRScore("count(*)", "(select id_person from person where academic_degree='кн' or academic_degree='дн')")) * 300).ToString()
                + Environment.NewLine + Environment.NewLine + distinctColumn["text"];

            distinctColumn = ScienceLeaderCount("магистрант");
            text_table_3_4_4_7 = Convert.ToInt32(float.Parse(distinctColumn["count"]) / float.Parse(calcNIRScore("count(*)", "(select id_person from person)")) * 200).ToString()
                + Environment.NewLine + Environment.NewLine + distinctColumn["text"];
        }

        //ППС в НИР
        private void NirPPS(ref string text_table_3_2_1, ref string text_table_3_2_2, ref int score_table_3_2_2)
        {
            text_table_3_2_1 = Convert.ToInt32(float.Parse(calcNIRScore("count(*)", "(select distinct id_person from reasearch_work where nir=true)")) / float.Parse(PPS_count) * 100).ToString();

            distinctColumn = calcAllNir("distinct title_reasearch", "reasearch_work", "nir=true");
            if (distinctColumn["text"] != "")
            {
                text_table_3_2_2 = distinctColumn["text"] + Environment.NewLine + Environment.NewLine;
                score_table_3_2_2 = int.Parse(distinctColumn["count"]) * 100;
            }
            distinctColumn = calcAllNir("distinct title_reasearch", "reasearch_work", "nir=false");
            if (distinctColumn["text"] != "")
            {
                text_table_3_2_2 += distinctColumn["text"];
                score_table_3_2_2 += int.Parse(distinctColumn["count"]) * 20;
            }
            text_table_3_2_2 = score_table_3_2_2.ToString() + Environment.NewLine + Environment.NewLine + text_table_3_2_2;
        }

        //Для учебников, статей, конференций ППС конвертация
        private Dictionary<string, string> ConvertDictBooks(Dictionary<string, string> dictInput)
        {
            Dictionary<string, string> dictOutput = new Dictionary<string, string>();
            insertText = "";

            foreach (KeyValuePair<string, string> keyValue in dictInput)
            {
                if (keyValue.Key == "count")
                {
                    if (keyValue.Value == "0")
                    {
                        dictOutput.Add("count", "");
                    }
                    else
                    {
                        dictOutput.Add("count", keyValue.Value);
                    }
                    continue;
                }
                insertText += keyValue.Value + Environment.NewLine + keyValue.Key + Environment.NewLine + Environment.NewLine;
            }
            dictOutput.Add("text", insertText);
            return dictOutput;
        }

        //Для изобретений, удельный вес конвертация
        private Dictionary<string, string> ConvertDictInventions(Dictionary<string, string> dictInput)
        {
            Dictionary<string, string> dictOutput = new Dictionary<string, string>();
            insertText = "";

            foreach (KeyValuePair<string, string> keyValue in dictInput)
            {
                if (keyValue.Key == "count")
                {
                    if (keyValue.Value == "0")
                    {
                        dictOutput.Add("count", "");
                    }
                    else
                    {
                        dictOutput.Add("count", keyValue.Value);
                    }
                    continue;
                }
                insertText += keyValue.Value + Environment.NewLine;
            }
            dictOutput.Add("text", insertText);
            return dictOutput;
        }

        //Для курсант. конф. конвертация 
        private Dictionary<string, string> ConvertDictConfKurs(Dictionary<string, string> dictInput)
        {
            Dictionary<string, string> dictOutput = new Dictionary<string, string>();
            insertText = "";
            foreach (KeyValuePair<string, string> keyValue in dictInput)
            {
                if (keyValue.Key == "count")
                {
                    if (keyValue.Value == "0")
                    {
                        dictOutput.Add("count", "");
                    }
                    else
                    {
                        dictOutput.Add("count", keyValue.Value);
                    }
                    continue;
                }
                insertText += keyValue.Value.Substring(0, keyValue.Value.Length - 2) + Environment.NewLine + keyValue.Key + Environment.NewLine;
            }

            dictOutput.Add("text", insertText);
            return dictOutput;
        }

        private Dictionary<string, string> ConvertDictConfThesis(Dictionary<string, string> dictInput)
        {

            Dictionary<string, string> dictOutput = new Dictionary<string, string>();
            insertText = "";
            foreach(KeyValuePair<string, string> keyValue in dictInput)
            {
                if (keyValue.Key == "count")
                {
                    if (keyValue.Value == "0")
                    {
                        dictOutput.Add("count", "");
                    }
                    else
                    {
                        dictOutput.Add("count", keyValue.Value);
                    }
                    continue;
                }
                insertText += keyValue.Key + " -" + keyValue.Value + Environment.NewLine + Environment.NewLine;
            }

            dictOutput.Add("text", insertText);
            return dictOutput;
        }

        //Учебники
        private void Textbooks(ref string text_table_3_2_3_1, ref string text_table_3_2_3_2, ref int score_table_3_2_3_1, ref int score_table_3_2_3_2)
        {
            distinctColumn = ConvertDictBooks(BooksCalc("Учебник"));
            if (distinctColumn["text"] != "")
            {
                text_table_3_2_3_1 = distinctColumn["text"] + Environment.NewLine;
                score_table_3_2_3_1 = Convert.ToInt32(float.Parse(distinctColumn["count"]) / float.Parse(PPS_count) * 200);
            }
            text_table_3_2_3_1 = score_table_3_2_3_1.ToString() + Environment.NewLine + Environment.NewLine + text_table_3_2_3_1;

            distinctColumn = ConvertDictBooks(BooksCalc("Учебное пособие"));
            if (distinctColumn["text"] != "")
            {
                text_table_3_2_3_2 = distinctColumn["text"] + Environment.NewLine;
                score_table_3_2_3_2 = Convert.ToInt32(float.Parse(distinctColumn["count"]) / float.Parse(PPS_count) * 150);
            }
            text_table_3_2_3_2 = score_table_3_2_3_2.ToString() + Environment.NewLine + Environment.NewLine + text_table_3_2_3_2;
        }

        private void button7_Click(object sender, EventArgs e)
        {
        }

        //Изобретателькая и рац. работа
        private void PPSInvention(ref string text_table_3_2_4_1, ref string text_table_3_2_4_2, ref int score_table_3_2_4_1, ref int score_table_3_2_4_2)
        {
            distinctColumn = ConvertDictInventions(calcInventions("view_invention='изобретение' and status_invention='поданная заявка'"));
            if (distinctColumn["text"] != "")
            {
                text_table_3_2_4_1 = distinctColumn["text"] + "поданная заявка на изобретение" + Environment.NewLine + Environment.NewLine;
                score_table_3_2_4_1 = int.Parse(distinctColumn["count"]) * 20;
            }

            distinctColumn = ConvertDictInventions(calcInventions("view_invention='полезная модель' and status_invention='поданная заявка'"));
            if (distinctColumn["text"] != "")
            {
                text_table_3_2_4_1 += distinctColumn["text"] + "поданная заявка на полезную модель" + Environment.NewLine + Environment.NewLine;
                score_table_3_2_4_1 += int.Parse(distinctColumn["count"]) * 10;
            }

            distinctColumn = ConvertDictInventions(calcInventions("view_invention='изобретение' and status_invention='положительный ответ'"));
            if (distinctColumn["text"] != "")
            {
                text_table_3_2_4_1 += distinctColumn["text"] +  "положительный ответ на изобретение" + Environment.NewLine + Environment.NewLine;
                score_table_3_2_4_1 += int.Parse(distinctColumn["count"]) * 50;
            }

            distinctColumn = ConvertDictInventions(calcInventions("view_invention='полезная модель' and status_invention='положительный ответ'"));
            if (distinctColumn["text"] != "")
            {
                text_table_3_2_4_1 += distinctColumn["text"] +  "положительный ответ на полезную модель" + Environment.NewLine + Environment.NewLine;
                score_table_3_2_4_1 += int.Parse(distinctColumn["count"]) * 30;
            }
            text_table_3_2_4_1 = score_table_3_2_4_1.ToString() + Environment.NewLine + Environment.NewLine + text_table_3_2_4_1;

            distinctColumn = ConvertDictInventions(calcInventions("view_invention='рац. предложение' and status_invention='принято'"));
            if (distinctColumn["text"] != "")
            {
                text_table_3_2_4_2 = distinctColumn["text"] + Environment.NewLine + "Рац. предложение" + Environment.NewLine;
                score_table_3_2_4_2 = int.Parse(distinctColumn["count"]) * 10;
            }
            text_table_3_2_4_2 = score_table_3_2_4_2.ToString() + Environment.NewLine + Environment.NewLine + text_table_3_2_4_2;

        }

        //Монографии и статьи
        private void MonografyArticles(ref string text_table_3_3_3_1, ref string text_table_3_3_3_2, ref string text_table_3_3_3_3, ref int score_table_3_3_3_1, ref int score_table_3_3_3_2, ref int score_table_3_3_3_3)
        {
            distinctColumn = ConvertDictBooks(BooksCalc("Монография"));
            if (distinctColumn["text"] != "")
            {
                text_table_3_3_3_1 = distinctColumn["text"] + Environment.NewLine;
                score_table_3_3_3_1 = Convert.ToInt32(float.Parse(distinctColumn["count"]) / float.Parse(PPS_count) * 300);
            }
            text_table_3_3_3_1 = score_table_3_3_3_1.ToString() + Environment.NewLine + Environment.NewLine + text_table_3_3_3_1;

            distinctColumn = ConvertDictBooks(CalcPpsArticles("Рецензируемая"));
            if (distinctColumn["text"] != "")
            {
                text_table_3_3_3_2 = distinctColumn["text"] + Environment.NewLine;
                score_table_3_3_3_2 = Convert.ToInt32(float.Parse(distinctColumn["count"]) / float.Parse(PPS_count) * 300);
            }
            text_table_3_3_3_2 = score_table_3_3_3_2.ToString() + Environment.NewLine + Environment.NewLine + text_table_3_3_3_2;

            distinctColumn = ConvertDictBooks(CalcPpsArticles("Нерецензируемая"));
            if (distinctColumn["text"] != "")
            {
                text_table_3_3_3_3 = distinctColumn["text"] + Environment.NewLine;
                score_table_3_3_3_3 = Convert.ToInt32(float.Parse(distinctColumn["count"]) / float.Parse(PPS_count) * 100);
            }
            text_table_3_3_3_3 = score_table_3_3_3_3.ToString() + Environment.NewLine + Environment.NewLine + text_table_3_3_3_3;
        }

        //Выставки ППС
        private void ExhibitiobsPPS(
             ref string text_table_3_3_2_1,
             ref string text_table_3_3_2_2,
             ref string text_table_3_3_2_3,
             ref string text_table_3_3_2_4,
             ref string text_table_3_3_2_5,
             ref string text_table_3_3_2_6,
             ref string text_table_3_3_2_7,
             ref string text_table_3_3_2_8,
             ref string text_table_3_3_2_9,
             ref string text_table_3_3_2_10,
             ref string text_table_3_3_2_11,
             ref string text_table_3_3_2_12
            )
        {
            text_table_3_3_2_1 = calcPpsExhibitions("distinct title_exhibitions", "exhibitions", "status_exhibitions='Вузовский' and diplom_status='1'", 300);
            text_table_3_3_2_2 = calcPpsExhibitions("distinct title_exhibitions", "exhibitions", "status_exhibitions='Вузовский' and diplom_status='2'", 250);
            text_table_3_3_2_3 = calcPpsExhibitions("distinct title_exhibitions", "exhibitions", "status_exhibitions='Вузовский' and diplom_status='3'", 200);
            text_table_3_3_2_4 = calcPpsExhibitions("distinct title_exhibitions", "exhibitions", "status_exhibitions='Вузовский' and diplom_status='0'", 100);
            text_table_3_3_2_5 = calcPpsExhibitions("distinct title_exhibitions", "exhibitions", "status_exhibitions='Республиканский' and diplom_status='1'", 400);
            text_table_3_3_2_6 = calcPpsExhibitions("distinct title_exhibitions", "exhibitions", "status_exhibitions='Республиканский' and diplom_status='2'", 350);
            text_table_3_3_2_7 = calcPpsExhibitions("distinct title_exhibitions", "exhibitions", "status_exhibitions='Республиканский' and diplom_status='3'", 300);
            text_table_3_3_2_8 = calcPpsExhibitions("distinct title_exhibitions", "exhibitions", "status_exhibitions='Республиканский' and diplom_status='0'", 200);
            text_table_3_3_2_9 = calcPpsExhibitions("distinct title_exhibitions", "exhibitions", "status_exhibitions='Международный' and diplom_status='1'", 500);
            text_table_3_3_2_10 = calcPpsExhibitions("distinct title_exhibitions", "exhibitions", "status_exhibitions='Республиканский' and diplom_status='2'", 450);
            text_table_3_3_2_11 = calcPpsExhibitions("distinct title_exhibitions", "exhibitions", "status_exhibitions='Республиканский' and diplom_status='3'", 400);
            text_table_3_3_2_12 = calcPpsExhibitions("distinct title_exhibitions", "exhibitions", "status_exhibitions='Республиканский' and diplom_status='0'", 300);

        }

        //Конференции ППС
        private void ConfPPS(ref string text_table_3_3_1_1, ref string text_table_3_3_1_3, ref string text_table_3_3_1_5, ref int score_table_3_3_1_1, ref int score_table_3_3_1_3, ref int score_table_3_3_1_5)
        {
            distinctColumn = ConvertDictBooks(calcPpsConf("Вузовский"));
            if (distinctColumn["text"] != "")
            {
                text_table_3_3_1_1 = distinctColumn["text"];
                score_table_3_3_1_1 = Convert.ToInt32(float.Parse(distinctColumn["count"]) / float.Parse(PPS_count) * 100);
            }
            text_table_3_3_1_1 = score_table_3_3_1_1.ToString() + Environment.NewLine + Environment.NewLine + text_table_3_3_1_1;

            distinctColumn = ConvertDictBooks(calcPpsConf("Республиканский"));
            if (distinctColumn["text"] != "")
            {
                text_table_3_3_1_3 = distinctColumn["text"];
                score_table_3_3_1_3 = Convert.ToInt32(float.Parse(distinctColumn["count"]) / float.Parse(PPS_count) * 200);
            }
            text_table_3_3_1_3 = score_table_3_3_1_3.ToString() + Environment.NewLine + Environment.NewLine + text_table_3_3_1_3;

            distinctColumn = ConvertDictBooks(calcPpsConf("Международный"));
            if (distinctColumn["text"] != "")
            {
                text_table_3_3_1_5 = distinctColumn["text"];
                score_table_3_3_1_5 = Convert.ToInt32(float.Parse(distinctColumn["count"]) / float.Parse(PPS_count) * 300);
            }
            text_table_3_3_1_5 = score_table_3_3_1_5.ToString() + Environment.NewLine + Environment.NewLine + text_table_3_3_1_5;
        }

        //Конференции курсантов
        private void ConfKursants(ref string text_table_3_3_1_2, ref string text_table_3_3_1_4, ref string text_table_3_3_1_6, ref int score_table_3_3_1_2, ref int score_table_3_3_1_4, ref int score_table_3_3_1_6)
        {
            distinctColumn = ConvertDictConfKurs(calcKursantConf("Вузовский"));
            if (distinctColumn["text"] != "")
            {
                text_table_3_3_1_2 = distinctColumn["text"];
                score_table_3_3_1_2 = Convert.ToInt32(float.Parse(distinctColumn["count"]) / float.Parse(PPS_count) * 0.1 * 100);
            }
            text_table_3_3_1_2 = score_table_3_3_1_2.ToString() + Environment.NewLine + Environment.NewLine + text_table_3_3_1_2;

            distinctColumn = ConvertDictConfKurs(calcKursantConf("Республиканский"));
            if (distinctColumn["text"] != "")
            {
                text_table_3_3_1_4 = distinctColumn["text"];
                score_table_3_3_1_4 = Convert.ToInt32(float.Parse(distinctColumn["count"]) / float.Parse(PPS_count) * 0.2 * 100);
            }
            text_table_3_3_1_4 = score_table_3_3_1_4.ToString() + Environment.NewLine + Environment.NewLine + text_table_3_3_1_4;

            distinctColumn = ConvertDictConfKurs(calcKursantConf("Международный"));
            if (distinctColumn["text"] != "")
            {
                text_table_3_3_1_6 = distinctColumn["text"];
                score_table_3_3_1_6 = Convert.ToInt32(float.Parse(distinctColumn["count"]) / float.Parse(PPS_count) * 0.5 * 100);
            }
            text_table_3_3_1_6 = score_table_3_3_1_6.ToString() + Environment.NewLine + Environment.NewLine + text_table_3_3_1_6;

        }


        //Под-ка науч раб
        private void ScienceEmpl(ref string text_table_3_4_1_1, ref string text_table_3_4_1_2, ref string text_table_3_4_1_3, ref string text_table_3_4_1_4, ref string text_table_3_4_1_5,
            ref int score_table_3_4_1_1, ref int score_table_3_4_1_2, ref int score_table_3_4_1_3, ref int score_table_3_4_1_4, ref int score_table_3_4_1_5)
        {
            distinctColumn = ConvertDictInventions(CalcPpsWeight("Name, Fam, mid_name", "person", "academic_degree='ДН'"));
            if (distinctColumn["text"] != "")
            {
                text_table_3_4_1_1 = distinctColumn["text"] + Environment.NewLine;
                score_table_3_4_1_1 = Convert.ToInt32(float.Parse(distinctColumn["count"]) / float.Parse(PPS_count) * 1000);
            }
            text_table_3_4_1_1 = score_table_3_4_1_1.ToString() + Environment.NewLine + Environment.NewLine + text_table_3_4_1_1;

            distinctColumn = ConvertDictInventions(CalcPpsWeight("Name, Fam, mid_name", "person", "scholarship='адьюнкт (очная форма)'"));
            if (distinctColumn["text"] != "")
            {
                text_table_3_4_1_2 = distinctColumn["text"] + Environment.NewLine;
                score_table_3_4_1_2 = Convert.ToInt32(float.Parse(distinctColumn["count"]) / float.Parse(PPS_count) * 700);
            }
            text_table_3_4_1_2 = score_table_3_4_1_2.ToString() + Environment.NewLine + Environment.NewLine + text_table_3_4_1_2;

            distinctColumn = ConvertDictInventions(CalcPpsWeight("Name, Fam, mid_name", "person", "scholarship='адьюнкт (заочная форма)'"));
            if (distinctColumn["text"] != "")
            {
                text_table_3_4_1_3 = distinctColumn["text"] + Environment.NewLine;
                score_table_3_4_1_3 = Convert.ToInt32(float.Parse(distinctColumn["count"]) / float.Parse(PPS_count) * 600);
            }
            text_table_3_4_1_3 = score_table_3_4_1_3.ToString() + Environment.NewLine + Environment.NewLine + text_table_3_4_1_3;

            distinctColumn = ConvertDictInventions(CalcPpsWeight("Name, Fam, mid_name", "person", "scholarship='соискатель ДН'"));
            if (distinctColumn["text"] != "")
            {
                text_table_3_4_1_4 = distinctColumn["text"] + Environment.NewLine;
                score_table_3_4_1_4 = Convert.ToInt32(float.Parse(distinctColumn["count"]) / float.Parse(PPS_count) * 900);
            }
            text_table_3_4_1_4 = score_table_3_4_1_4.ToString() + Environment.NewLine + Environment.NewLine + text_table_3_4_1_4;

            distinctColumn = ConvertDictInventions(CalcPpsWeight("Name, Fam, mid_name", "person", "scholarship='соискатель КН'"));
            if (distinctColumn["text"] != "")
            {
                text_table_3_4_1_5 = distinctColumn["text"] + Environment.NewLine;
                score_table_3_4_1_5 = Convert.ToInt32(float.Parse(distinctColumn["count"]) / float.Parse(PPS_count) * 500);
            }
            text_table_3_4_1_5 = score_table_3_4_1_5.ToString() + Environment.NewLine + Environment.NewLine + text_table_3_4_1_5;
        }

        //Отзывы учебники
        private void TextbooksReview(ref string text_table_3_2_5, ref int score_table_3_2_5)
        {
            distinctColumn = calcAllReviews("distinct title_review, person", "reviews", "tag='Документ/Учебник/Пособие'");
            if (distinctColumn.ContainsKey("text") && distinctColumn["text"] != "")
            {
                text_table_3_2_5 = distinctColumn["text"];
                score_table_3_2_5 = Convert.ToInt32(float.Parse(distinctColumn["count"]) / float.Parse(PPS_count) * 500);
            }
            text_table_3_2_5 = score_table_3_2_5.ToString() + Environment.NewLine + Environment.NewLine + text_table_3_2_5;
        }


        //Тезисы
        private void Tesis(ref string text_table_3_3_3_4, ref int score_table_3_3_3_4)
        {
            distinctColumn = ConvertDictConfThesis(CalcThesisConf());
            if (distinctColumn["text"] != "")
            {
                text_table_3_3_3_4 = distinctColumn["text"] + Environment.NewLine;
                score_table_3_3_3_4 = Convert.ToInt32(float.Parse(distinctColumn["count"]) / float.Parse(PPS_count) * 50);
            }
            text_table_3_3_3_4 = score_table_3_3_3_4.ToString() + Environment.NewLine + Environment.NewLine + text_table_3_3_3_4;
        }

        //Количество ппс высш. квалиф.
        private void CountSciencePPS(ref string text_table_3_4_3_1, ref string text_table_3_4_3_2, ref string text_table_3_4_3_3, ref string text_table_3_4_3_4,
            ref string text_table_3_4_3_5, ref string text_table_3_4_3_6, ref string text_table_3_4_3_7, ref string text_table_3_4_3_8,
            ref int score_table_3_4_3_5, ref int score_table_3_4_3_6, ref int score_table_3_4_3_7, ref int score_table_3_4_3_8)
        {
            int score = Convert.ToInt32(float.Parse(calcNIRScore("count(*)", "person where academic_degree='КН' or academic_degree='ДН'")) / int.Parse(PPS_count) * 100);

            if (score > 40)
            {
                text_table_3_4_3_1 = "50";
            }
            else if (score > 20 && score < 40)
            {
                text_table_3_4_3_2 = "40";
            }
            else if (score > 0 && score < 20)
            {
                text_table_3_4_3_3 = "20";
            }

            distinctColumn = CalcPpsWeight("Fam, Name, mid_name", "person", "academic_title='профессор'");
            if (distinctColumn["FIO"] != "")
            {
                text_table_3_4_3_5 = distinctColumn["FIO"] + Environment.NewLine;
                score_table_3_4_3_5 = Convert.ToInt32(float.Parse(distinctColumn["count"]) / float.Parse(calcNIRScore("professor", "ShDS")) * 100);
            }
            text_table_3_4_3_5 = score_table_3_4_3_5.ToString() + Environment.NewLine + Environment.NewLine + text_table_3_4_3_5;

            distinctColumn = distinctColumn = CalcPpsWeight("Fam, Name, mid_name", "person", "academic_title='доцент'");
            if (distinctColumn["FIO"] != "")
            {
                text_table_3_4_3_6 = distinctColumn["FIO"] + Environment.NewLine;
                score_table_3_4_3_6 = Convert.ToInt32(float.Parse(distinctColumn["count"]) / float.Parse(calcNIRScore("assistant_professor", "ShDS")) * 100);
            }
            text_table_3_4_3_6 = score_table_3_4_3_6.ToString() + Environment.NewLine + Environment.NewLine + text_table_3_4_3_6;

            distinctColumn = CalcPpsWeight("Fam, Name, mid_name", "person", "academic_degree='ДН'");
            if (distinctColumn["FIO"] != "")
            {
                text_table_3_4_3_7 = distinctColumn["FIO"] + Environment.NewLine;
                score_table_3_4_3_7 = Convert.ToInt32(float.Parse(distinctColumn["count"]) / float.Parse(PPS_count) * 1000);
            }
            text_table_3_4_3_7 = score_table_3_4_3_7.ToString() + Environment.NewLine + Environment.NewLine + text_table_3_4_3_7;

            distinctColumn = CalcPpsWeight("Fam, Name, mid_name", "person", "academic_degree='КН'");
            if (distinctColumn["FIO"] != "")
            {
                text_table_3_4_3_8 = distinctColumn["FIO"] + Environment.NewLine;
                score_table_3_4_3_8 = Convert.ToInt32(float.Parse(distinctColumn["count"]) / float.Parse(PPS_count) * 500);
            }
            text_table_3_4_3_8 = score_table_3_4_3_8.ToString() + Environment.NewLine + Environment.NewLine + text_table_3_4_3_8;
        }

        //Отзывы
        private void AllReviews(ref string text_table_3_4_4_1, ref string text_table_3_4_4_2, ref string text_table_3_4_4_3, ref string text_table_3_4_4_4,
            ref int score_table_3_4_4_1, ref int score_table_3_4_4_2, ref int score_table_3_4_4_3, ref int score_table_3_4_4_4)
        {
            distinctColumn = calcAllReviews("distinct title_review,person", "reviews", "tag='доксторской диссертации' and token='Экспертиза'");
            if (distinctColumn["text"] != "")
            {
                text_table_3_4_4_1 = distinctColumn["text"] + Environment.NewLine;
                score_table_3_4_4_1 = Convert.ToInt32(float.Parse(distinctColumn["count"]) / float.Parse(calcNIRScore("count(*)", "person where academic_degree='ДН'")) * 250);
            }
            text_table_3_4_4_1 = score_table_3_4_4_1.ToString() + Environment.NewLine + Environment.NewLine + text_table_3_4_4_1;

            distinctColumn = calcAllReviews("distinct title_review,person", "reviews", "tag='кандидантской диссертации' and token='Экспертиза'");
            if (distinctColumn["text"] != "")
            {
                text_table_3_4_4_2 = distinctColumn["text"] + Environment.NewLine;
                score_table_3_4_4_2 = Convert.ToInt32(float.Parse(distinctColumn["count"]) / float.Parse(calcNIRScore("count(*)", "person where academic_degree='КН' or academic_degree='КН'")) * 200);
            }
            text_table_3_4_4_2 = score_table_3_4_4_2.ToString() + Environment.NewLine + Environment.NewLine + text_table_3_4_4_2;

            distinctColumn = calcAllReviews("distinct title_review,person", "reviews", "tag='доксторской диссертации' and token='Отзыв'");
            if (distinctColumn["text"] != "")
            {
                text_table_3_4_4_3 = distinctColumn["text"] + Environment.NewLine;
                score_table_3_4_4_3 = Convert.ToInt32(float.Parse(distinctColumn["count"]) / float.Parse(calcNIRScore("count(*)", "person where academic_degree='ДН'")) * 300);
            }
            text_table_3_4_4_3 = score_table_3_4_4_3.ToString() + Environment.NewLine + Environment.NewLine + text_table_3_4_4_3;

            distinctColumn = calcAllReviews("distinct title_review,person", "reviews", "tag='кандидантской диссертации' and token='Отзыв'");
            if (distinctColumn["text"] != "")
            {
                text_table_3_4_4_4 = distinctColumn["text"] + Environment.NewLine;
                score_table_3_4_4_4 = Convert.ToInt32(float.Parse(distinctColumn["count"]) / float.Parse(calcNIRScore("count(*)", "person where academic_degree='КН' or academic_degree='КН'")) * 200);
            }
            text_table_3_4_4_4 = score_table_3_4_4_4.ToString() + Environment.NewLine + Environment.NewLine + text_table_3_4_4_4;
        }

        private void btn_exit_2_Click(object sender, EventArgs e)
        {
            System.Environment.Exit(0); // Приложение завершается и возвращает ОС указанное параметром значение
        }

        private void btn_restore_Click(object sender, EventArgs e)
        {
            if (this.WindowState == FormWindowState.Maximized)
            {
                this.WindowState = FormWindowState.Normal;
            }
            else
                this.WindowState = FormWindowState.Maximized;
        }

        private void btn_minimize_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        private void Form_reports_MouseDown(object sender, MouseEventArgs e)
        {
            base.Capture = false;
            Message m = Message.Create(base.Handle, 0xa1, new IntPtr(2), IntPtr.Zero);
            this.WndProc(ref m);
        }

        private void checkBox14_CheckedChanged(object sender, EventArgs e)
        {

            if (all_cheaked)
            {
                checkBox1.Checked = false;
                checkBox2.Checked = false;
                checkBox3.Checked = false;
                checkBox4.Checked = false;
                checkBox5.Checked = false;
                checkBox6.Checked = false;
                checkBox7.Checked = false;
                checkBox8.Checked = false;
                checkBox9.Checked = false;
                checkBox10.Checked = false;
                checkBox11.Checked = false;
                checkBox12.Checked = false;
                checkBox13.Checked = false;
            }
            else
            {
                checkBox1.Checked = true;
                checkBox2.Checked = true;
                checkBox3.Checked = true;
                checkBox4.Checked = true;
                checkBox5.Checked = true;
                checkBox6.Checked = true;
                checkBox7.Checked = true;
                checkBox8.Checked = true;
                checkBox9.Checked = true;
                checkBox10.Checked = true;
                checkBox11.Checked = true;
                checkBox12.Checked = true;
                checkBox13.Checked = true;
            }
            all_cheaked = !all_cheaked;
        }

        private void Form_reports_Load(object sender, EventArgs e)
        {
        }

        //Вставка данных в 4-ую таблицу
        private void insert_data_table4(Word.Table table)
        {

            Dictionary<string[], string> dictInput =  ConfInTable_2();
            foreach (KeyValuePair<string[], string> keyValue in dictInput)
            {
                table.Rows.Add(table.Rows[2]);
                table.Cell(2, 1).Range.Text = keyValue.Key[0];
                table.Cell(2, 2).Range.Text = keyValue.Key[1] + Environment.NewLine + keyValue.Key[2];
                table.Cell(2, 3).Range.Text = keyValue.Value;
                table.Cell(2, 4).Range.Text = keyValue.Key[3];
            }
        }

        //Вставка данных в 5-ую таблицу
        private void insert_data_table5(Word.Table table)
        {

            Dictionary<string[], string> dictInput = Textbooks_Articles_2();
            foreach (KeyValuePair<string[], string> keyValue in dictInput)
            {
                table.Rows.Add(table.Rows[2]);
                table.Cell(2, 1).Range.Text = keyValue.Key[0];
                table.Cell(2, 2).Range.Text = keyValue.Value + "//" + Environment.NewLine + keyValue.Key[1];
            }
        }


        private Dictionary<string[], string> Textbooks_Articles_2()
        {
            try
            {
                connection.Open();
                Dictionary<string[], string> distinctColumn = new Dictionary<string[], string>();

                insertText = "";

                command.CommandText = "select title_textbook,title_edition from (select  distinct title_textbook,title_edition from textbooks);";
                reader = command.ExecuteReader();

                while (reader.Read())
                {
                    command2.CommandText = "select person.Fam,person.Name,person.mid_name from textbooks inner join person on textbooks.id_person = person.id_person where title_textbook='" + reader[0].ToString() + "' and title_edition='" + reader[1].ToString() + "';";
                    reader2 = command2.ExecuteReader();

                    string[] keyDict = { reader[0].ToString(), reader[1].ToString() };

                    if (!distinctColumn.ContainsKey(keyDict))
                        distinctColumn.Add(keyDict, "");

                    while (reader2.Read())
                    {
                        insertText += reader2[0].ToString() + " " + reader2[1].ToString()[0] + "." + reader2[2].ToString()[0] + "., ";
                    }

                    distinctColumn[keyDict] += insertText.Substring(0, insertText.Length - 2);
                    insertText = "";
                    if (!reader2.IsClosed)
                        reader2.Close();
                }
                connection.Close();
                reader2.Close();
                reader.Close();

                return distinctColumn;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                connection.Close();
                connection.Dispose();
                return null;
            }
        }

        //Для документа квартального отчёта таблица конференции
        private Dictionary<string[], string> ConfInTable_2()
        {
            try
            {
                connection.Open();
                Dictionary< string[], string> distinctColumn = new Dictionary<string[], string>();


                countRow = 0;
                insertText = "";
                command.CommandText = "select distinct name_conference,conference_list.id_record_conf,place_conference,date_conf,results from conference inner join conference_list on conference_list.id_record_conf= conference.id_record_conf;";
                reader = command.ExecuteReader();

                while (reader.Read())
                {
                    countRow++;
                    command2.CommandText = "select distinct id_person from conference inner join conference_list on conference.id_record_conf = conference_list.id_record_conf where  conference_list.id_record_conf=" + reader[1].ToString() + ";";
                    string[] keyDict = { reader[0].ToString(), reader[2].ToString(), Convert.ToDateTime(reader[3]).Date.ToString().Substring(0, 10), reader[4].ToString() };
                    reader2 = command2.ExecuteReader();
                    if (!distinctColumn.ContainsKey(keyDict))
                        distinctColumn.Add(keyDict, "");

                    while (reader2.Read())
                    {
                        command3.CommandText = " select Fam, Name, mid_name  from  person  where id_person=" + reader2[0].ToString() + ";";
                        reader3 = command3.ExecuteReader();
                        while (reader3.Read())
                        {
                            insertText += reader3[0].ToString() + " " + reader3[1].ToString()[0] + "." + reader3[2].ToString()[0] + "., ";
                        }
                        if (!reader3.IsClosed)
                            reader3.Close();
                    }
                    distinctColumn[keyDict] += insertText.Substring(0, insertText.Length - 2);

                    insertText = "";
                    if (!reader2.IsClosed)
                        reader2.Close();

                }
                connection.Close();
                reader2.Close();
                reader.Close();

                return distinctColumn;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                connection.Close();
                connection.Dispose();
                return null;
            }
        }

        //Для сумма первой таблицы главного отчёта
        private float sum_table1(Word.Table table)
        {
            float sum = 0;
            float number = 0;

            List<string> allRows = new List<string>();
            List<string[]> allCountInRows = new List<string[]>();

            for (int i = 1; i <= table.Rows.Count - 1; i++)
            {
                try
                {
                    allRows.Add(table.Cell(i, 5).Range.Text.Replace("\r", " ").Replace("\a", " ").Replace("\n", " "));
                }
                catch (Exception ex)
                {
                    allRows.Add(table.Cell(i, 4).Range.Text.Replace("\r", " ").Replace("\a", " ").Replace("\n", " "));
                }
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
                            sum += number;
                        }
                    }
                }
                else
                {
                    if (float.TryParse(allCountInRows[i][0], out number))
                    {
                        sum += number;
                    }
                }
            }

            return sum;
        }

        //Сумма для общего отчёта
        private float sum_table(Word.Table table)
        {
            float sum = 0;
            float number = 0;

            List<string> allRows = new List<string>();
            List<string[]> allCountInRows = new List<string[]>();

            for (int i = 1; i <= table.Rows.Count - 1; i++)
            {
                try
                {
                    allRows.Add(table.Cell(i, 5).Range.Text.Replace("\r", " ").Replace("\a", " ").Replace("\n", " "));
                }
                catch (Exception ex)
                {
                    allRows.Add(table.Cell(i, 4).Range.Text.Replace("\r", " ").Replace("\a", " ").Replace("\n", " "));
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
                    sum += number;
                }
            }

            return sum;
        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        //Для показа процесса формирования отчёта
        private void show_info(string text, Color color, int time)
        {

            Task.Run(() =>
            {
                BeginInvoke(new MethodInvoker(delegate
                {
                    label_file.ForeColor = color;
                    label_file.Text = text;
                    label_file.Visible = true;
                }));

                Thread.Sleep(time);

                BeginInvoke(new MethodInvoker(delegate
                {
                    label_file.Visible = false;
                }));
            });
        }

        //Для показа процесса формирования отчёта
        private void show_info_2(string text, Color color, int time)
        {

            Task.Run(() =>
            {
                BeginInvoke(new MethodInvoker(delegate
                {
                    label1.ForeColor = color;
                    label1.Text = text;
                    label1.Visible = true;
                }));

                Thread.Sleep(time);

                BeginInvoke(new MethodInvoker(delegate
                {
                    label1.Visible = false;
                }));
            });
        }

        private async void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Word Documents|*.doc";
            openFileDialog.Title = "Выберите шаблон отчёта";
            openFileDialog.ShowDialog();

            if (openFileDialog.FileName != "")
            {
                pathToTemplateFile = openFileDialog.FileName;
            }
            else
            {
                MessageBox.Show("Вы не выбрали файл шаблона. Попробуем найти его на обычном место", "Упс!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }


            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.Filter = "Word Documents|*.doc";
            saveFileDialog1.Title = "Сохраните отчёт";
            saveFileDialog1.ShowDialog();
            string filename = "";

            if (saveFileDialog1.FileName != "")
            {
                filename = saveFileDialog1.FileName;
                label_file.ForeColor = Color.Gold;
                label_file.Text = "Формирование отчёта...";
                label_file.Visible = true;
            }
            else
            {
                MessageBox.Show("Вы не сохранили файл. Сформируйте отчёт ещё раз", "Упс!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            await Task.Run(() =>
            {
                if (checkBox1.Checked)
                    ConfAndSemPPS(
                       ref text_table_3_1_2,
                       ref score_table_3_1_2);
                if (checkBox2.Checked)
                    NirPPS(
                        ref text_table_3_2_1,
                        ref text_table_3_2_2,
                        ref score_table_3_2_2);
                if (checkBox13.Checked)
                    AllReviews(
                        ref text_table_3_4_4_1,
                        ref text_table_3_4_4_2,
                        ref text_table_3_4_4_3,
                        ref text_table_3_4_4_4,
                        ref score_table_3_4_4_1,
                        ref score_table_3_4_4_2,
                        ref score_table_3_4_4_3,
                        ref score_table_3_4_4_4);
                if (checkBox3.Checked)
                    Textbooks(
                        ref text_table_3_2_3_1,
                        ref text_table_3_2_3_2,
                        ref score_table_3_2_3_1,
                        ref score_table_3_2_3_2);
                if (checkBox4.Checked)
                    PPSInvention(
                        ref text_table_3_2_4_1,
                        ref text_table_3_2_4_2,
                        ref score_table_3_2_4_1,
                        ref score_table_3_2_4_2);
                if (checkBox9.Checked)
                    MonografyArticles(
                        ref text_table_3_3_3_1,
                        ref text_table_3_3_3_2,
                        ref text_table_3_3_3_3,
                        ref score_table_3_3_3_1,
                        ref score_table_3_3_3_2,
                        ref score_table_3_3_3_3);
                if (checkBox8.Checked)
                    ExhibitiobsPPS(
                        ref text_table_3_3_2_1,
                        ref text_table_3_3_2_2,
                        ref text_table_3_3_2_3,
                        ref text_table_3_3_2_4,
                        ref text_table_3_3_2_5,
                        ref text_table_3_3_2_6,
                        ref text_table_3_3_2_7,
                        ref text_table_3_3_2_8,
                        ref text_table_3_3_2_9,
                        ref text_table_3_3_2_10,
                        ref text_table_3_3_2_11,
                        ref text_table_3_3_2_12);
                if (checkBox6.Checked)
                    ConfPPS(
                        ref text_table_3_3_1_1,
                        ref text_table_3_3_1_3,
                        ref text_table_3_3_1_5,
                        ref score_table_3_3_1_1,
                        ref score_table_3_3_1_3,
                        ref score_table_3_3_1_5);
                if (checkBox7.Checked)
                    ConfKursants(
                        ref text_table_3_3_1_2,
                        ref text_table_3_3_1_4,
                        ref text_table_3_3_1_6,
                        ref score_table_3_3_1_2,
                        ref score_table_3_3_1_4,
                        ref score_table_3_3_1_6);
                if (checkBox11.Checked)
                    ScienceEmpl(
                        ref text_table_3_4_1_1,
                        ref text_table_3_4_1_2,
                        ref text_table_3_4_1_3,
                        ref text_table_3_4_1_4,
                        ref text_table_3_4_1_5,
                        ref score_table_3_4_1_1,
                        ref score_table_3_4_1_2,
                        ref score_table_3_4_1_3,
                        ref score_table_3_4_1_4,
                        ref score_table_3_4_1_5);
                if (checkBox5.Checked)
                    TextbooksReview(
                        ref text_table_3_2_5,
                        ref score_table_3_2_5);
                if (checkBox10.Checked)
                    Tesis(
                        ref text_table_3_3_3_4,
                        ref score_table_3_3_3_4);
                if (checkBox12.Checked)
                {
                    CountSciencePPS(
                        ref text_table_3_4_3_1,
                        ref text_table_3_4_3_2,
                        ref text_table_3_4_3_3,
                        ref text_table_3_4_3_4,
                        ref text_table_3_4_3_5,
                        ref text_table_3_4_3_6,
                        ref text_table_3_4_3_7,
                        ref text_table_3_4_3_8,
                        ref score_table_3_4_3_5,
                        ref score_table_3_4_3_6,
                        ref score_table_3_4_3_7,
                        ref score_table_3_4_3_8);
                    ScienceLeader(ref text_table_3_4_4_5, ref text_table_3_4_4_6, ref text_table_3_4_4_7);
                }

            });

            var wordApp = new Word.Application();
            wordApp.Visible = false;

            try
            {
                await Task.Run(() =>
                {

                    var wordDocument = wordApp.Documents.Open(pathToTemplateFile);

                    ReplaceWordStub("{table3.1_2}", text_table_3_1_2, wordDocument);
                    ReplaceWordStub("{table3.2_1}", text_table_3_2_1, wordDocument);
                    ReplaceWordStub("{table3.2_2}", text_table_3_2_2, wordDocument);
                    ReplaceWordStub("{table3.2_3_1}", text_table_3_2_3_1, wordDocument);
                    ReplaceWordStub("{table3.2_3_2}", text_table_3_2_3_2, wordDocument);
                    ReplaceWordStub("{table3.2_4_1}", text_table_3_2_4_1, wordDocument);
                    ReplaceWordStub("{table3.2_4_2}", text_table_3_2_4_2, wordDocument);
                    ReplaceWordStub("{table3.2_5}", text_table_3_2_5, wordDocument);
                    ReplaceWordStub("{table3.3_1_1}", text_table_3_3_1_1, wordDocument);
                    ReplaceWordStub("{table3.3_1_2}", text_table_3_3_1_2, wordDocument);
                    ReplaceWordStub("{table3.3_1_3}", text_table_3_3_1_3, wordDocument);
                    ReplaceWordStub("{table3.3_1_4}", text_table_3_3_1_4, wordDocument);
                    ReplaceWordStub("{table3.3_1_5}", text_table_3_3_1_5, wordDocument);
                    ReplaceWordStub("{table3.3_1_6}", text_table_3_3_1_6, wordDocument);
                    ReplaceWordStub("{table3.3_2_1}", text_table_3_3_2_1, wordDocument);
                    ReplaceWordStub("{table3.3_2_2}", text_table_3_3_2_2, wordDocument);
                    ReplaceWordStub("{table3.3_2_3}", text_table_3_3_2_3, wordDocument);
                    ReplaceWordStub("{table3.3_2_4}", text_table_3_3_2_4, wordDocument);
                    ReplaceWordStub("{table3.3_2_5}", text_table_3_3_2_5, wordDocument);
                    ReplaceWordStub("{table3.3_2_6}", text_table_3_3_2_6, wordDocument);
                    ReplaceWordStub("{table3.3_2_7}", text_table_3_3_2_7, wordDocument);
                    ReplaceWordStub("{table3.3_2_8}", text_table_3_3_2_8, wordDocument);
                    ReplaceWordStub("{table3.3_2_9}", text_table_3_3_2_9, wordDocument);
                    ReplaceWordStub("{table3.3_2_10}", text_table_3_3_2_10, wordDocument);
                    ReplaceWordStub("{table3.3_2_11}", text_table_3_3_2_11, wordDocument);
                    ReplaceWordStub("{table3.3_2_12}", text_table_3_3_2_12, wordDocument);
                    ReplaceWordStub("{table3.3_3_1}", text_table_3_3_3_1, wordDocument);
                    ReplaceWordStub("{table3.3_3_2}", text_table_3_3_3_2, wordDocument);
                    ReplaceWordStub("{table3.3_3_3}", text_table_3_3_3_3, wordDocument);
                    ReplaceWordStub("{table3.3_3_4}", text_table_3_3_3_4, wordDocument);
                    ReplaceWordStub("{table3.4_1_1}", text_table_3_4_1_1, wordDocument);
                    ReplaceWordStub("{table3.4_1_2}", text_table_3_4_1_2, wordDocument);
                    ReplaceWordStub("{table3.4_1_3}", text_table_3_4_1_3, wordDocument);
                    ReplaceWordStub("{table3.4_1_4}", text_table_3_4_1_4, wordDocument);
                    ReplaceWordStub("{table3.4_1_5}", text_table_3_4_1_5, wordDocument);
                    ReplaceWordStub("{table3.4_3_1}", text_table_3_4_3_1, wordDocument);
                    ReplaceWordStub("{table3.4_3_2}", text_table_3_4_3_2, wordDocument);
                    ReplaceWordStub("{table3.4_3_3}", text_table_3_4_3_3, wordDocument);
                    ReplaceWordStub("{table3.4_3_4}", text_table_3_4_3_4, wordDocument);
                    ReplaceWordStub("{table3.4_3_5}", text_table_3_4_3_5, wordDocument);
                    ReplaceWordStub("{table3.4_3_6}", text_table_3_4_3_6, wordDocument);
                    ReplaceWordStub("{table3.4_3_7}", text_table_3_4_3_7, wordDocument);
                    ReplaceWordStub("{table3.4_3_8}", text_table_3_4_3_8, wordDocument);
                    ReplaceWordStub("{table3.4_4_1}", text_table_3_4_4_1, wordDocument);
                    ReplaceWordStub("{table3.4_4_2}", text_table_3_4_4_2, wordDocument);
                    ReplaceWordStub("{table3.4_4_3}", text_table_3_4_4_3, wordDocument);
                    ReplaceWordStub("{table3.4_4_4}", text_table_3_4_4_4, wordDocument);
                    ReplaceWordStub("{table3.4_4_5}", text_table_3_4_4_5, wordDocument);
                    ReplaceWordStub("{table3.4_4_6}", text_table_3_4_4_6, wordDocument);
                    ReplaceWordStub("{table3.4_4_7}", text_table_3_4_4_7, wordDocument);

                    wordDocument.SaveAs(filename);
                    wordDocument.Close();

                    var wordDocumentSum = wordApp.Documents.Open(filename);

                    wordDocumentSum.Activate();

                    Word.Table table1 = wordDocumentSum.Tables[1];
                    Word.Table table2 = wordDocumentSum.Tables[2];
                    Word.Table table3 = wordDocumentSum.Tables[3];
                    Word.Table table4 = wordDocumentSum.Tables[4];

                    float summa_table1 = sum_table1(table1);
                    float summa_table2 = sum_table(table2);
                    float summa_table3 = sum_table(table3);
                    float summa_table4 = sum_table(table4);

                    ReplaceWordStub("{table3.1_sum}", summa_table1.ToString(), wordDocumentSum);
                    ReplaceWordStub("{table3.2_sum}", summa_table2.ToString(), wordDocumentSum);
                    ReplaceWordStub("{table3.3_sum}", summa_table3.ToString(), wordDocumentSum);
                    ReplaceWordStub("{table3.4_sum}", summa_table4.ToString(), wordDocumentSum);

                    ReplaceWordStub("{table3.1_sum_1}", summa_table1.ToString(), wordDocumentSum);
                    ReplaceWordStub("{table3.2_sum_1}", summa_table2.ToString(), wordDocumentSum);
                    ReplaceWordStub("{table3.3_sum_1}", summa_table3.ToString(), wordDocumentSum);
                    ReplaceWordStub("{table3.4_sum_1}", summa_table4.ToString(), wordDocumentSum);
                    ReplaceWordStub("{sum}", (summa_table1 * 0.15 + summa_table2 * 0.3 + summa_table3 * 0.2 + summa_table4 * 0.35).ToString(), wordDocumentSum);

                    wordDocumentSum.SaveAs(filename);
                    wordDocumentSum.Close();

                    show_info("Отчёт сохранён", Color.Green, 2000);
                });

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                text_table_3_1_2 = "";
                text_table_3_2_1 = "";
                text_table_3_2_2 = "";
                text_table_3_2_3_1 = "";
                text_table_3_2_3_2 = "";
                text_table_3_2_4_1 = "";
                text_table_3_2_4_2 = "";

                text_table_3_2_5 = "";

                text_table_3_3_1_1 = "";
                text_table_3_3_1_2 = "";
                text_table_3_3_1_3 = "";
                text_table_3_3_1_4 = "";
                text_table_3_3_1_5 = "";
                text_table_3_3_1_6 = "";

                text_table_3_3_2_1 = "";
                text_table_3_3_2_2 = "";
                text_table_3_3_2_3 = "";
                text_table_3_3_2_4 = "";
                text_table_3_3_2_5 = "";
                text_table_3_3_2_6 = "";
                text_table_3_3_2_7 = "";
                text_table_3_3_2_8 = "";
                text_table_3_3_2_9 = "";
                text_table_3_3_2_10 = "";
                text_table_3_3_2_11 = "";
                text_table_3_3_2_12 = "";

                text_table_3_3_3_1 = "";
                text_table_3_3_3_2 = "";
                text_table_3_3_3_3 = "";
                text_table_3_3_3_4 = "";

                text_table_3_4_1_1 = "";
                text_table_3_4_1_2 = "";
                text_table_3_4_1_3 = "";
                text_table_3_4_1_4 = "";
                text_table_3_4_1_5 = "";

                text_table_3_4_3_1 = "";
                text_table_3_4_3_2 = "";
                text_table_3_4_3_3 = "";
                text_table_3_4_3_4 = "";
                text_table_3_4_3_5 = "";
                text_table_3_4_3_6 = "";
                text_table_3_4_3_7 = "";
                text_table_3_4_3_8 = "";

                text_table_3_4_4_1 = "";
                text_table_3_4_4_2 = "";
                text_table_3_4_4_3 = "";
                text_table_3_4_4_4 = "";

                score_table_3_1_2 = 0;

                score_table_3_2_4_2 = 0;
                
                score_table_3_2_2 = 0;
                score_table_3_2_3_1 = 0;
                score_table_3_2_3_2 = 0;
                score_table_3_2_4_1 = 0;
                score_table_3_2_5 = 0;
                
                score_table_3_3_1_1 = 0;
                score_table_3_3_1_2 = 0;
                score_table_3_3_1_3 = 0;
                score_table_3_3_1_4 = 0;
                score_table_3_3_1_5 = 0;
                score_table_3_3_1_6 = 0;
                
                score_table_3_3_3_1 = 0;
                score_table_3_3_3_2 = 0;
                score_table_3_3_3_3 = 0;
                score_table_3_3_3_4 = 0;
                
                score_table_3_4_1_1 = 0;
                score_table_3_4_1_2 = 0;
                score_table_3_4_1_3 = 0;
                score_table_3_4_1_4 = 0;
                score_table_3_4_1_5 = 0;
                
                score_table_3_4_3_5 = 0;
                score_table_3_4_3_6 = 0;
                score_table_3_4_3_7 = 0;
                score_table_3_4_3_8 = 0;
                
                score_table_3_4_4_1 = 0;
                score_table_3_4_4_2 = 0;
                score_table_3_4_4_3 = 0;
                score_table_3_4_4_4 = 0;

                wordApp.Quit();
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void btn_back_Click(object sender, EventArgs e)
        {
            this.Hide();
            Form_Admin form_admin = new Form_Admin();
            form_admin.ShowDialog();
        }

        private void checkBox15_CheckedChanged(object sender, EventArgs e)
        {
            if (all_cheaked_2)
            {
                checkBox21.Checked = false;
                checkBox22.Checked = false;
                checkBox23.Checked = false;
                checkBox25.Checked = false;
                checkBox24.Checked = false;
                checkBox26.Checked = false;
                checkBox25.Checked = false;
                checkBox27.Checked = false;
                checkBox28.Checked = false;
            }
            else
            {
                checkBox21.Checked = true;
                checkBox22.Checked = true;
                checkBox23.Checked = true;
                checkBox25.Checked = true;
                checkBox24.Checked = true;
                checkBox26.Checked = true;
                checkBox25.Checked = true;
                checkBox27.Checked = true;
                checkBox28.Checked = true;
            }
            all_cheaked_2 = !all_cheaked_2;
        }

        private async void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Word Documents|*.doc";
            openFileDialog.Title = "Выберите шаблон отчёта за квартал";
            openFileDialog.ShowDialog();

            if (openFileDialog.FileName != "")
            {
                pathToTemplFile = openFileDialog.FileName;
            }
            else
            {
                MessageBox.Show("Вы не выбрали файл шаблона. Попробуем найти его на обычном место", "Упс!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }


            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.Filter = "Word Documents|*.doc";
            saveFileDialog1.Title = "Сохраните отчёт";
            saveFileDialog1.ShowDialog();
            string filename = "";

            if (saveFileDialog1.FileName != "")
            {
                filename = saveFileDialog1.FileName;
                label1.ForeColor = Color.Gold;
                label1.Text = "Формирование отчёта...";
                label1.Visible = true;
            }
            else
            {
                MessageBox.Show("Вы не сохранили файл. Сформируйте отчёт ещё раз", "Упс!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }


            var wordApp = new Word.Application();
            wordApp.Visible = false;

            await Task.Run(() =>
            {
                try 
                {
                    var wordDocument = wordApp.Documents.Open(pathToTemplFile);
                    Word.Table table4 = wordDocument.Tables[4];
                    Word.Table table5 = wordDocument.Tables[5];

                    if(checkBox28.Checked)
                        insert_data_table4(table4);
                    if (checkBox24.Checked)
                        insert_data_table5(table5);

                    wordDocument.SaveAs(filename);
                    wordDocument.Close();

                    show_info_2("Отчёт сохранён", Color.Green, 2000);

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                finally
                {
                    wordApp.Quit();
                }


                
            });
        }

        private void checkBox12_CheckedChanged(object sender, EventArgs e)
        {

        }
    }
}
