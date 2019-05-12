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
    public partial class Form_textbooks : Form
    {

        //Данные пользователя
        Form_login login_form;

        //Имя таблица, для внесения данных
        private string table_name = "";


        //флаг для определения хочет ли пользователь обновить строку
        private bool flag_update = false;

        //для отслеживание выбранной статьи
        private string id_textbook = "0";

        //Поля для поиска
        private string search_field = "";

        private OleDbConnection connection = new OleDbConnection();
        public Form_textbooks()
        {
            InitializeComponent();
        }

        public Form_textbooks(Form_login form_login, string table_name)
        {
            InitializeComponent();
            login_form = form_login;
            this.table_name = table_name;
            connection.ConnectionString = Form_login.connectString;

            label_user_name.Text = login_form.fam_person + " " + login_form.name_person + " " + login_form.mid_name_person;
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


                    command.CommandText = "delete from " + table_name + " where id_textbook=" + int.Parse(id_textbook) + ";";



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
                        command.CommandText = "update " + table_name + " set title_textbook='" + txt_textbook.Text + "'," + "title_edition='" + txt_edition.Text + "'," + "date_stamp='" +
                            dateTimePicker.Value.Date + "'," + "number_sheets='" + txt_sheets.Text + "',neck='" + comboBox.SelectedItem.ToString() +"',book_view='" + comboBox_book.SelectedItem.ToString() + "' where id_textbook=" + int.Parse(id_textbook) + ";";
                    }
                    else
                    {
                        command.CommandText = "insert into " + table_name + " (title_textbook,title_edition,date_stamp,number_sheets,neck,book_view,id_person) values('" + txt_textbook.Text + "','" +
                            txt_edition.Text + "','" + dateTimePicker.Value.Date + "'," + int.Parse(txt_sheets.Text) + ", '" + comboBox.SelectedItem.ToString() + "','"+ comboBox_book.SelectedItem.ToString() + "'," + login_form.id_person + ");";
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

        private void btn_back_Click(object sender, EventArgs e)
        {
            this.Hide();
            Form_selection_table form_table = new Form_selection_table(login_form);
            form_table.ShowDialog();
        }

        private void Form_textbooks_FormClosed(object sender, FormClosedEventArgs e)
        {
            System.Environment.Exit(0); // Приложение завершается и возвращает ОС указанное параметром значение

        }


        private void load_data_table()
        {
            //Получаем все данные из таблицы для соответствующего пользователя
            try
            {
                connection.Open();
                OleDbCommand command = new OleDbCommand();
                command.Connection = connection;
                command.CommandText = "select id_textbook,title_textbook,title_edition,date_stamp,number_sheets,neck,book_view from " + table_name + " where id_person=" + login_form.id_person + ";";


                OleDbDataAdapter adapter = new OleDbDataAdapter(command);
                DataTable table = new DataTable();

                Task.Run(() =>
                {
                    BeginInvoke(new MethodInvoker(delegate
                    {
                        adapter.Fill(table);
                        dataGridView.DataSource = table;

                        dataGridView.Columns[0].HeaderText = "№";
                        dataGridView.Columns[1].HeaderText = "Название учебника";
                        dataGridView.Columns[2].HeaderText = "Название издания";
                        dataGridView.Columns[3].HeaderText = "Дата печати";
                        dataGridView.Columns[4].HeaderText = "Количество страниц";
                        dataGridView.Columns[5].HeaderText = "Гриф";
                        dataGridView.Columns[6].HeaderText = "Вид";
                    }));
                });


                dataGridView.DataSource = table;

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


        private void clean_input()
        {
            txt_edition.Text = "";
            txt_sheets.Text = "";
            txt_textbook.Text = "";
            comboBox.Text = "";
            comboBox_book.Text = "";
        }
        private void Form_textbooks_Load(object sender, EventArgs e)
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
                id_textbook = row.Cells[0].Value.ToString();
                txt_textbook.Text = row.Cells[1].Value.ToString();
                txt_edition.Text = row.Cells[2].Value.ToString();
                dateTimePicker.Value = DateTime.Parse(row.Cells[3].Value.ToString());
                txt_sheets.Text = row.Cells[4].Value.ToString();
                comboBox.SelectedItem = row.Cells[5].Value.ToString();
                comboBox_book.SelectedItem = row.Cells[6].Value.ToString();

            }
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

        private void Form_textbooks_MouseDown(object sender, MouseEventArgs e)
        {
            base.Capture = false;
            Message m = Message.Create(base.Handle, 0xa1, new IntPtr(2), IntPtr.Zero);
            this.WndProc(ref m);
        }

        private void comboBox_search_SelectedIndexChanged(object sender, EventArgs e)
        {
           
           txt_search.Text = "";

            switch (comboBox_search.SelectedItem.ToString())
            {
                case "Название учебника":
                    search_field = "title_textbook";
                    break;
                case "Название издания":
                    search_field = "title_edition";
                    break;
                case "Количество страниц":
                    search_field = "number_sheets";
                    break;
                case "Гриф":
                    search_field = "neck";
                    break;
                case "Вид":
                    search_field = "book_view";
                    break;
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            comboBox_search.Enabled = !comboBox_search.Enabled;
            txt_search.Enabled = !txt_search.Enabled;
            load_data_table();
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
                    command.CommandText = "select id_textbook,title_textbook,title_edition,date_stamp,number_sheets,neck,book_view from " + table_name + " where " + search_field + " LIKE '%" + txt_search.Text + "%' and id_person=" + login_form.id_person + ";";


                    OleDbDataAdapter adapter = new OleDbDataAdapter(command);
                    DataTable table = new DataTable();

                    Task.Run(() =>
                    {
                        BeginInvoke(new MethodInvoker(delegate
                        {
                            adapter.Fill(table);
                            dataGridView.DataSource = table;

                            dataGridView.Columns[0].HeaderText = "№";
                            dataGridView.Columns[1].HeaderText = "Название учебника";
                            dataGridView.Columns[2].HeaderText = "Название издания";
                            dataGridView.Columns[3].HeaderText = "Дата печати";
                            dataGridView.Columns[4].HeaderText = "Количество страниц";
                            dataGridView.Columns[5].HeaderText = "Гриф";
                            dataGridView.Columns[6].HeaderText = "Вид";
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

        private void panel3_Paint(object sender, PaintEventArgs e)
        {

        }

        private void label_book_Click(object sender, EventArgs e)
        {

        }
    }
}
