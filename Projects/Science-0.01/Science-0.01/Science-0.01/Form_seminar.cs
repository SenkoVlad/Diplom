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
    public partial class Form_seminar : Form
    {
        //Данные пользователя
        Form_login login_form;


        //Имя таблица, для внесения данных
        private string table_name = "";

        //Для чтения таблицы журнала конф
        private int id_record_conf = 0;

        //флаг для определения хочет ли пользователь обновить строку
        private bool flag_update = false;

        //для отслеживание выбранной конференции
        private string id_seminar = "0";

        //Поля для поиска
        private string search_field = "";

        //Для чтения данных из таблицы
        OleDbDataReader reader = null;
        private OleDbConnection connection = new OleDbConnection();

        public Form_seminar()
        {
            InitializeComponent();
            connection.ConnectionString = Form_login.connectString;
        }

        public Form_seminar(Form_login form_login, string table_name)
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
                    command.CommandText = "select id_record_conf from conference_list where name_conference='" + comboBox_name_sem.SelectedItem.ToString() + "' and place_conference='" + comboBox_place_sem.SelectedItem.ToString() + "';";
                    reader = command.ExecuteReader();
                    id_record_conf = 0;
                    while (reader.Read())
                    {
                        id_record_conf = int.Parse(reader[0].ToString());
                    }
                    if (id_record_conf == 0)
                    {
                        MessageBox.Show("Выберите правильное соответствие места и конференции. Скорей всего в данном месте нету такой конференции. Обратитесь к администратору", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        connection.Close();
                        throw new Exception("Попробуйте ещё раз");
                    }
                    reader.Close();
                    if (flag_update)
                    {
                        command.CommandText = "update " + table_name + " set title_report='" + txt_report.Text + "'," + "date_seminar='" +
                            dateTimePicker.Value.Date + "'," + "id_record_conf=" + id_record_conf + ",number_sheets=" + int.Parse(txt_sheets.Text) + "," + "results='" + txt_results.Text + "', status='" + comboBox.SelectedItem.ToString() + "' where id_seminar=" + int.Parse(id_seminar) + ";";
                    }
                    else
                    {
                        command.CommandText = "insert into " + table_name + " (title_report,date_seminar,id_record_conf,status,number_sheets,results,id_person) values('" + txt_report.Text + "','" +
                            dateTimePicker.Value.Date + "'," + id_record_conf + ",'" + comboBox.SelectedItem.ToString() + "', " + int.Parse(txt_sheets.Text) + ",'" + txt_results.Text + "', " + login_form.id_person + ");";
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


                    command.CommandText = "delete from " + table_name + " where id_senimar=" + int.Parse(id_seminar) + ";";



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
                id_seminar = row.Cells[0].Value.ToString();
                txt_report.Text = row.Cells[1].Value.ToString();
                comboBox_name_sem.SelectedItem = row.Cells[2].Value.ToString();
                dateTimePicker.Value = DateTime.Parse(row.Cells[3].Value.ToString());
                comboBox_place_sem.SelectedItem = row.Cells[4].Value.ToString();
                comboBox.SelectedItem = row.Cells[5].Value.ToString();
                txt_sheets.Text = row.Cells[6].Value.ToString();
                txt_results.Text = row.Cells[7].Value.ToString();

            }
        }

        private void clean_input()
        {
            comboBox_name_sem.Text = "";
            comboBox_place_sem.Text = "";
            txt_report.Text = "";
            txt_results.Text = "";
            txt_sheets.Text = "";
            comboBox.Text = "";
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

        //Вывести данные в DataGrid
        private  void load_data_table()
        {
            //Получаем все данные из таблицы для соответствующего пользователя
            try
            {
                connection.Open();
                OleDbCommand command = new OleDbCommand();
                command.Connection = connection;
                command.CommandText = "select name_conference,place_conference from conference_list;";
                reader = command.ExecuteReader();

                comboBox_name_sem.Items.Clear();
                comboBox_place_sem.Items.Clear();

                while (reader.Read())
                {
                    if (!comboBox_name_sem.Items.Contains(reader[0].ToString()))
                        comboBox_name_sem.Items.Add(reader[0].ToString());
                    if (!comboBox_place_sem.Items.Contains(reader[1].ToString()))
                        comboBox_place_sem.Items.Add(reader[1].ToString());
                }
                reader.Close();

                command.CommandText = "select id_seminar,title_report,name_conference,date_seminar,place_conference,status,number_sheets,results from " + table_name + " inner join conference_list on conference_list.id_record_conf = seminar.id_record_conf where id_person=" + login_form.id_person + ";";


                OleDbDataAdapter adapter = new OleDbDataAdapter(command);
                DataTable table = new DataTable();

                Task.Run(() =>
                {
                    BeginInvoke(new MethodInvoker(delegate
                    {
                        adapter.Fill(table);
                        dataGridView.DataSource = table;

                        dataGridView.Columns[0].HeaderText = "№";
                        dataGridView.Columns[1].HeaderText = "Название доклада";
                        dataGridView.Columns[2].HeaderText = "Название семинара";
                        dataGridView.Columns[3].HeaderText = "Дата";
                        dataGridView.Columns[4].HeaderText = "Место";
                        dataGridView.Columns[5].HeaderText = "Статус";
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

        private void Form_seminar_FormClosed(object sender, FormClosedEventArgs e)
        {
                System.Environment.Exit(0); // Приложение завершается и возвращает ОС указанное параметром значение
        }

        private void Form_seminar_Load(object sender, EventArgs e)
        {
            load_data_table();
        }

        private void txt_sheets_TextChanged(object sender, EventArgs e)
        {

        }

        private void txt_results_TextChanged(object sender, EventArgs e)
        {

        }

        private void txt_seminar_TextChanged(object sender, EventArgs e)
        {

        }

        private void txt_report_TextChanged(object sender, EventArgs e)
        {

        }

        private void Form_seminar_MouseDown(object sender, MouseEventArgs e)
        {
            base.Capture = false;
            Message m = Message.Create(base.Handle, 0xa1, new IntPtr(2), IntPtr.Zero);
            this.WndProc(ref m);
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
                    command.CommandText = "select id_seminar,title_report,name_conference,date_seminar,place_conference,status,number_sheets,results from " + table_name + " inner join conference_list on conference_list.id_record_conf = seminar.id_record_conf where " + search_field + " LIKE '%" + txt_search.Text + "%' and id_person=" + login_form.id_person + ";";


                    OleDbDataAdapter adapter = new OleDbDataAdapter(command);
                    DataTable table = new DataTable();

                    Task.Run(() =>
                    {
                        BeginInvoke(new MethodInvoker(delegate
                        {
                            adapter.Fill(table);
                            dataGridView.DataSource = table;

                            dataGridView.Columns[0].HeaderText = "№";
                            dataGridView.Columns[1].HeaderText = "Название доклада";
                            dataGridView.Columns[2].HeaderText = "Название семинара";
                            dataGridView.Columns[3].HeaderText = "Дата";
                            dataGridView.Columns[4].HeaderText = "Место";
                            dataGridView.Columns[5].HeaderText = "Статус";
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
                case "Название доклада":
                    search_field = "title_report";
                    break;
                case "Название семинара":
                    search_field = "name_conference";
                    break;
                case "Место":
                    search_field = "place_conference";
                    break;
                case "Статус":
                    search_field = "status";
                    break;
                case "Количество страниц":
                    search_field = "number_sheets";
                    break;
                case "Результаты":
                    search_field = "results";
                    break;
            }
        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void btn_exit_2_Click_1(object sender, EventArgs e)
        {
            System.Environment.Exit(0); // Приложение завершается и возвращает ОС указанное параметром значение
        }

        private void btn_restore_Click_1(object sender, EventArgs e)
        {
            if (this.WindowState == FormWindowState.Maximized)
            {
                this.WindowState = FormWindowState.Normal;
            }
            else
                this.WindowState = FormWindowState.Maximized;
        }

        private void btn_minimize_Click_1(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }
    }
}
