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
    public partial class Form_review : Form
    {

        //Данные пользователя
        Form_login login_form;


        //Имя таблица, для внесения данных
        private string table_name = "";

        //флаг для определения хочет ли пользователь обновить строку
        private bool flag_update = false;

        //для отслеживание выбранной конференции
        private string id_review = "0";

        //Поля для поиска
        private string search_field = "";

        private OleDbConnection connection = new OleDbConnection();

        //Для чтения данных из таблицы
        OleDbDataReader reader = null;

        public Form_review()
        {
            InitializeComponent();
            connection.ConnectionString = Form_login.connectString;
        }

        public Form_review(Form_login form_login, string table_name)
        {
            InitializeComponent();
            login_form = form_login;
            this.table_name = table_name;
            connection.ConnectionString = Form_login.connectString;

            label_user_name.Text = login_form.fam_person + " " + login_form.name_person + " " + login_form.mid_name_person;
        }
        private void btn_back_Click(object sender, EventArgs e)
        {
            this.Hide();
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
                        command.CommandText = "update " + table_name + " set title_review='" + txt_name.Text + "'," + "person='" + comboBox_persons.SelectedItem.ToString() + "'," + "date_review='" +
                            dateTimePicker.Value.Date + "'," + "results='" + txt_results.Text + "',number_sheets=" + int.Parse(txt_sheets.Text) + "," + "token='" + comboBox_token.SelectedItem.ToString() + "', tag='" + comboBox_tag.SelectedItem.ToString() + "' where id_review=" + int.Parse(id_review) + ";";
                    }
                    else
                    {
                        command.CommandText = "insert into " + table_name + " (title_review,person,token,tag,date_review,number_sheets,results,id_person) values('" + txt_name.Text + "','" +
                            comboBox_persons.SelectedItem.ToString() + "','" + comboBox_token.SelectedItem.ToString() + "','" + comboBox_tag.SelectedItem.ToString() + "', '" + dateTimePicker.Value.Date + "'," + int.Parse(txt_sheets.Text) + ",'" + txt_results.Text + "', " + login_form.id_person + ");";
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

        private void Form_review_Load(object sender, EventArgs e)
        {
            load_data_table();
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


                    command.CommandText = "delete from " + table_name + " where id_review=" + int.Parse(id_review) + ";";



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

        private void btn_exit_Click(object sender, EventArgs e)
        {
            System.Environment.Exit(0); // Приложение завершается и возвращает ОС указанное параметром значение

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

        private void Form_review_FormClosed(object sender, FormClosedEventArgs e)
        {
            System.Environment.Exit(0); // Приложение завершается и возвращает ОС указанное параметром значение
        }


        private void clean_input()
        {
            txt_results.Text = "";
            txt_name.Text = "";
            comboBox_persons.Text = "";
            comboBox_tag.Text = "";
            comboBox_token.Text = "";
            txt_sheets.Text = "";
        }
        //Вывод информации по работе с таблицей асинхронно с "миганием"
        private void show_info(string text, Color color, int time)
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
                    label1.ForeColor = color;
                    label1.Text = text;
                    label1.Visible = false;
                }));
            });
        }

        //Вывести данные в DataGrid
        private  void load_data_table()
        {
            //Получаем все данные из таблицы для соответствующего пользователя
            try
            {
                connection.Open();
                OleDbCommand command = new OleDbCommand();
                string fam = "";
                string name = "";
                string mid_name = "";
                string full_name = "";
                command.Connection = connection;

                command.CommandText = "select Name,Fam,mid_name from person;";
                reader = command.ExecuteReader();

                comboBox_persons.Items.Clear();

                while (reader.Read())
                {
                    fam = reader[1].ToString();
                    name = reader[0].ToString();
                    mid_name = reader[2].ToString();
                    full_name = fam + " " + name + " " + mid_name;

                    if (!comboBox_persons.Items.Contains(full_name))
                        comboBox_persons.Items.Add(full_name);
                }
                reader.Close();

                command.CommandText = "select id_review,title_review,person,token,tag,date_review,number_sheets,results from " + table_name + " where id_person=" + login_form.id_person + ";";


                OleDbDataAdapter adapter = new OleDbDataAdapter(command);
                DataTable table = new DataTable();
                Task.Run(() =>
                {
                    BeginInvoke(new MethodInvoker(delegate
                    {
                        adapter.Fill(table);
                        dataGridView.DataSource = table;

                        dataGridView.Columns[0].HeaderText = "№";
                        dataGridView.Columns[1].HeaderText = "Название";
                        dataGridView.Columns[2].HeaderText = "Проводивший";
                        dataGridView.Columns[3].HeaderText = "Степень";
                        dataGridView.Columns[4].HeaderText = "Вид";
                        dataGridView.Columns[5].HeaderText = "Дата проведения";
                        dataGridView.Columns[6].HeaderText = "Количество страниц";
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
                id_review = row.Cells[0].Value.ToString();
                txt_name.Text = row.Cells[1].Value.ToString();
                comboBox_persons.SelectedItem = row.Cells[2].Value.ToString();
                comboBox_token.SelectedItem = row.Cells[3].Value.ToString();
                comboBox_tag.SelectedItem = row.Cells[4].Value.ToString();
                dateTimePicker.Value = DateTime.Parse(row.Cells[5].Value.ToString());
                txt_sheets.Text = row.Cells[6].Value.ToString();
                txt_results.Text = row.Cells[7].Value.ToString();
            }
        }

        private void Form_review_MouseDown(object sender, MouseEventArgs e)
        {
            base.Capture = false;
            Message m = Message.Create(base.Handle, 0xa1, new IntPtr(2), IntPtr.Zero);
            this.WndProc(ref m);
        }

        private void comboBox_token_SelectedIndexChanged(object sender, EventArgs e)
        {
            label_tag.Text = comboBox_token.SelectedItem.ToString() + " по";
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
                case "Название":
                    search_field = "title_review";
                    break;
                case "Проводившее лицо":
                    search_field = "person";
                    break;
                case "Количество листов":
                    search_field = "number_sheets";
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
                    command.CommandText = "select id_review,title_review,person,token,tag,date_review,number_sheets,results from " + table_name + " where " + search_field + " LIKE '%" + txt_search.Text + "%' and id_person=" + login_form.id_person + ";";


                    OleDbDataAdapter adapter = new OleDbDataAdapter(command);
                    DataTable table = new DataTable();

                    Task.Run(() =>
                    {
                        BeginInvoke(new MethodInvoker(delegate
                        {
                            adapter.Fill(table);
                            dataGridView.DataSource = table;

                            dataGridView.Columns[0].HeaderText = "№";
                            dataGridView.Columns[1].HeaderText = "Название";
                            dataGridView.Columns[2].HeaderText = "Проводивший";
                            dataGridView.Columns[3].HeaderText = "Степень";
                            dataGridView.Columns[4].HeaderText = "Вид";
                            dataGridView.Columns[5].HeaderText = "Дата проведения";
                            dataGridView.Columns[6].HeaderText = "Количество страниц";
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

        private void txt_person_TextChanged(object sender, EventArgs e)
        {

        }

    }
}
