using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Science_0._01
{
    public partial class Form_mill_training : Form
    {

        //Данные пользователя
        Form_login login_form;

        //Имя таблица, для внесения данных
        private string table_name = "";


        //флаг для определения хочет ли пользователь обновить строку
        private bool flag_update = false;

        //для отслеживание выбранных учений
        private string id_millitary_training = "0";

        //Поля для поиска
        private string search_field = "";

        private OleDbConnection connection = new OleDbConnection();

        public Form_mill_training()
        {
            InitializeComponent();
        }

        public Form_mill_training(Form_login form_login, string table_name)
        {
            InitializeComponent();
            login_form = form_login;
            this.table_name = table_name;
            connection.ConnectionString = Form_login.connectString;

            label_user_name.Text = login_form.fam_person + " " + login_form.name_person + " " + login_form.mid_name_person;
        }

        private void Form_mill_training_FormClosed(object sender, FormClosedEventArgs e)
        {
            System.Environment.Exit(0); // Приложение завершается и возвращает ОС указанное параметром значение
        }

        private void btn_exit_Click(object sender, EventArgs e)
        {
            System.Environment.Exit(0); // Приложение завершается и возвращает ОС указанное параметром значение
        }

        private void btn_back_Click(object sender, EventArgs e)
        {
            Form_selection_table form_table = new Form_selection_table(login_form);
            form_table.ShowDialog();
        }

        private async void btn_save_Click(object sender, EventArgs e)
        {
            if (panel.Enabled)
            {
                //Вносим новую строку
                try
                {
                    connection.Open();
                    OleDbCommand command = new OleDbCommand();
                    command.Connection = connection;
                    if (flag_update)
                    {
                        command.CommandText = "update " + table_name + " set title_training='" + txt_training.Text + "'," + "questions='" + txt_questions.Text + "'," + "date_training='" +
                            dateTimePicker.Value.Date + "'," + "place='" + txt_palce.Text + "',head='" + txt_head.Text + "',reporting='" + txt_reporting.Text + "',results='" + txt_results.Text+ "' where id_military_training=" + int.Parse(id_millitary_training) + ";";
                    }
                    else
                    {
                        command.CommandText = "insert into " + table_name + " (title_training,questions,date_training,place,head,reporting,results,id_person) values('" + txt_training.Text + "','" +
                            txt_questions.Text + "','" + dateTimePicker.Value.Date + "','" + txt_palce.Text + "', '" + txt_head.Text + "','" + txt_reporting.Text + "','" + txt_results.Text + "'," + login_form.id_person + ");";
                    }


                    await Task.Run(() => command.ExecuteNonQuery());

                    if (flag_update)
                        show_info("Данные обновлены", Color.Gold, 2000);
                    else
                        show_info("Данные внесены", Color.Green, 2000);

                    panel.Enabled = false;
                    flag_update = false;
                    connection.Close();

                    //Обновить таблицу
                    load_data_table();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    connection.Close();

                }
            }
            else
            {
                MessageBox.Show("Нажните на кнопку 'Новая'", "Упс!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void btn_new_Click(object sender, EventArgs e)
        {
            panel.Enabled = true;
            flag_update = false;
            clean_input();
        }

        private void btn_update_Click(object sender, EventArgs e)
        {
            panel.Enabled = true;
            flag_update = true;
        }

        private async void btn_delete_Click(object sender, EventArgs e)
        {

            DialogResult answer = MessageBox.Show(
            "Удалить запись?",
            "Сообщение",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Information,
            MessageBoxDefaultButton.Button1,
            MessageBoxOptions.DefaultDesktopOnly);

            if (answer == DialogResult.Yes)
            {
                try
                {
                    connection.Open();
                    OleDbCommand command = new OleDbCommand();
                    command.Connection = connection;


                    command.CommandText = "delete from " + table_name + " where id_military_training=" + int.Parse(id_millitary_training) + ";";



                    await Task.Run(() => command.ExecuteNonQuery());

                    show_info("Запись удалена", Color.Red, 2000);

                    panel.Enabled = false;
                    connection.Close();
                    //Обновить таблицу
                    load_data_table();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    connection.Close();
                }
            }
        }



        private void load_data_table()
        {
            //Получаем все данные из таблицы для соответствующего пользователя
            try
            {
                connection.Open();
                OleDbCommand command = new OleDbCommand();
                command.Connection = connection;
                command.CommandText = "select id_military_training,title_training,questions,date_training,place,head,reporting,results from " + table_name + " where id_person=" + login_form.id_person + ";";


                OleDbDataAdapter adapter = new OleDbDataAdapter(command);
                DataTable table = new DataTable();

                Task.Run(() =>
                {
                    BeginInvoke(new MethodInvoker(delegate
                    {
                        adapter.Fill(table);
                        dataGridView.DataSource = table;

                        dataGridView.Columns[0].HeaderText = "№";
                        dataGridView.Columns[1].HeaderText = "Название учений";
                        dataGridView.Columns[2].HeaderText = "Вопросы";
                        dataGridView.Columns[3].HeaderText = "Дата";
                        dataGridView.Columns[4].HeaderText = "Место";
                        dataGridView.Columns[5].HeaderText = "Руководитель";
                        dataGridView.Columns[6].HeaderText = "Форма отчётности";
                        dataGridView.Columns[7].HeaderText = "Результаты";

                    }));
                });

                connection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                connection.Close();

            }
        }

        //Вывод информации по работе с таблицей асинхронно с "миганием"
        private void show_info(string text, Color color, int time)
        {

            Task.Run(() =>
            {
                BeginInvoke(new MethodInvoker(delegate
                {
                    label_info.ForeColor = color;
                    label_info.Text = text;
                    label_info.Visible = true;
                }));

                Thread.Sleep(time);

                BeginInvoke(new MethodInvoker(delegate
                {
                    label_info.ForeColor = color;
                    label_info.Text = text;
                    label_info.Visible = false;
                }));
            });
        }

        private void Form_mill_training_Load(object sender, EventArgs e)
        {
            load_data_table();
        }

        private void dataGridView_SelectionChanged(object sender, EventArgs e)
        {
            panel.Enabled = false;

            DataGridViewCell cell = null;
            foreach (DataGridViewCell selectedCell in dataGridView.SelectedCells)
            {
                cell = selectedCell;
                break;
            }
            if (cell != null)
            {
                DataGridViewRow row = cell.OwningRow;
                id_millitary_training = row.Cells[0].Value.ToString();
                txt_training.Text = row.Cells[1].Value.ToString();
                txt_questions.Text = row.Cells[2].Value.ToString();
                dateTimePicker.Value = DateTime.Parse(row.Cells[3].Value.ToString());
                txt_palce.Text = row.Cells[4].Value.ToString();
                txt_head.Text = row.Cells[5].Value.ToString();
                txt_reporting.Text = row.Cells[6].Value.ToString();
                txt_results.Text = row.Cells[7].Value.ToString();
            }
        }

        private void clean_input()
        {
            txt_head.Text = "";
            txt_palce.Text = "";
            txt_questions.Text = "";
            txt_reporting.Text = "";
            txt_results.Text = "";
            txt_training.Text = "";
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

        private void Form_mill_training_MouseDown(object sender, MouseEventArgs e)
        {
            base.Capture = false;
            Message m = Message.Create(base.Handle, 0xa1, new IntPtr(2), IntPtr.Zero);
            this.WndProc(ref m);
        }

        private void panel3_Paint(object sender, PaintEventArgs e)
        {

        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            comboBox_search.Enabled = !comboBox_search.Enabled;
            txt_search.Enabled = !txt_search.Enabled;
            load_data_table();
        }

        private void comboBox_search_SelectedIndexChanged(object sender, EventArgs e)
        {

            txt_search.Text = "";

            switch (comboBox_search.SelectedItem.ToString())
            {
                case "Название учений":
                    search_field = "title_training";
                    break;
                case "Вопросы":
                    search_field = "questions";
                    break;
                case "Место":
                    search_field = "place";
                    break;
                case "Руководитель":
                    search_field = "head";
                    break;
                case "Форма отчётности":
                    search_field = "reporting";
                    break;
                case "Результаты":
                    search_field = "results";
                    break;
            }
        }

        private void txt_search_TextChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {


                //Получаем все данные из таблицы для соответствующего пользователя
                try
                {
                    connection.Open();
                    OleDbCommand command = new OleDbCommand();
                    command.Connection = connection;
                    command.CommandText = "select id_military_training,title_training,questions,date_training,place,head,reporting,results from " + table_name + " where " + search_field + " LIKE '%" + txt_search.Text + "%' and id_person=" + login_form.id_person + ";";


                    OleDbDataAdapter adapter = new OleDbDataAdapter(command);
                    DataTable table = new DataTable();

                    Task.Run(() =>
                    {
                        BeginInvoke(new MethodInvoker(delegate
                        {
                            adapter.Fill(table);
                            dataGridView.DataSource = table;

                            dataGridView.Columns[0].HeaderText = "№";
                            dataGridView.Columns[1].HeaderText = "Название учений";
                            dataGridView.Columns[2].HeaderText = "Вопросы";
                            dataGridView.Columns[3].HeaderText = "Дата";
                            dataGridView.Columns[4].HeaderText = "Место";
                            dataGridView.Columns[5].HeaderText = "Руководитель";
                            dataGridView.Columns[6].HeaderText = "Форма отчётности";
                            dataGridView.Columns[7].HeaderText = "Результаты";

                        }));
                    });

                    connection.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    connection.Close();

                }
            }
        }

        private void btn_exit_Click_1(object sender, EventArgs e)
        {
            System.Environment.Exit(0); // Приложение завершается и возвращает ОС указанное параметром значение
        }
    }
}
