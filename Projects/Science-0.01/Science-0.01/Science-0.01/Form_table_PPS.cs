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
    public partial class Form_table_PPS : Form
    {
        //Имя таблица, для внесения данных
        private string table_name = "person";

        //флаг для определения хочет ли пользователь обновить строку
        private bool flag_update = false;

        //для отслеживание выбранного пользователя
        private string id_person = "0";

        //Поля для поиска
        private string search_field = "";

        private OleDbConnection connection = new OleDbConnection();
        public Form_table_PPS()
        {
            InitializeComponent();
            connection.ConnectionString = Form_login.connectString;

            label_user_name.Text = "Admin";
        }



        private void panel3_Paint(object sender, PaintEventArgs e)
        {

        }

        private void btn_back_Click(object sender, EventArgs e)
        {
            this.Hide();
            Form_Admin form_table = new Form_Admin();
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
                        command.CommandText = "update " + table_name + " set Name='" + txt_name.Text + "'," + "Fam='" + txt_fam.Text + "'," + "mid_name='" +
                            txt_name2.Text + "'," + "scholarship='" + comboBox_level.SelectedItem.ToString()+ "'," + "academic_degree='" + comboBox_academic_degree.SelectedItem.ToString()  +"'," + "academic_title='" + comboBox_academic_title.SelectedItem.ToString()    + "' where id_person=" + int.Parse(id_person) + ";";
                    }
                    else
                    {
                        command.CommandText = "insert into " + table_name + " (Name,Fam,mid_name,scholarship,academic_degree,academic_title) values('" + txt_name.Text + "','" +
                            txt_fam.Text + "','" + txt_name2.Text + "','" + comboBox_level.SelectedItem.ToString() + "','" + comboBox_academic_degree.SelectedItem.ToString() + "','" + comboBox_academic_title.SelectedItem.ToString() + "');";
                    }


                    await Task.Run(() => command.ExecuteNonQuery());

                    if (flag_update)
                        show_info("Данные обновлены", Color.Gold, 2000);
                    else
                        show_info("Данные внесены", Color.GreenYellow, 2000);

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


                    command.CommandText = "delete from " + table_name + " where id_person=" + int.Parse(id_person) + ";";



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

        private void Form_table_PPS_FormClosed(object sender, FormClosedEventArgs e)
        {
            System.Environment.Exit(0); // Приложение завершается и возвращает ОС указанное параметром значение
        }

        private void comboBox_search_SelectedIndexChanged(object sender, EventArgs e)
        {
            txt_search.Text = "";
            switch (comboBox_search.SelectedItem.ToString())
            {
                case "Фамилия":
                    search_field = "Fam";
                    break;
                case "Имя":
                    search_field = "Name";
                    break;
                case "Отчество":
                    search_field = "mid_name";
                    break;
                case "Учёность":
                    search_field = "scholarship";
                    break;
                case "Учёная степень":
                    search_field = "academic_degree";
                    break;
                case "Учёное звание":
                    search_field = "academic_title";
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
                    command.CommandText = "select id_person,Name,Fam,mid_name,scholarship,academic_degree,academic_title from " + table_name + " where " + search_field + " LIKE '%" + txt_search.Text + "%';";


                    OleDbDataAdapter adapter = new OleDbDataAdapter(command);
                    DataTable table = new DataTable();

                    Task.Run(() =>
                    {
                        BeginInvoke(new MethodInvoker(delegate
                        {
                            adapter.Fill(table);
                            dataGridView.DataSource = table;

                            dataGridView.Columns[0].HeaderText = "№";
                            dataGridView.Columns[1].HeaderText = "Имя";
                            dataGridView.Columns[2].HeaderText = "Фамилия";
                            dataGridView.Columns[3].HeaderText = "Отчество";
                            dataGridView.Columns[4].HeaderText = "Учёность";
                            dataGridView.Columns[5].HeaderText = "Учёная степень";
                            dataGridView.Columns[6].HeaderText = "Учёное звание";

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
                id_person = row.Cells[0].Value.ToString();
                txt_name.Text = row.Cells[1].Value.ToString();
                txt_fam.Text = row.Cells[2].Value.ToString();
                txt_name2.Text = row.Cells[3].Value.ToString();
                comboBox_level.SelectedItem = row.Cells[4].Value.ToString();
                comboBox_academic_degree.SelectedItem = row.Cells[5].Value.ToString();
                comboBox_academic_title.SelectedItem = row.Cells[6].Value.ToString();

            }
        }

        private void btn_minimize_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
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

        private void Form_table_PPS_MouseDown(object sender, MouseEventArgs e)
        {
            base.Capture = false;
            Message m = Message.Create(base.Handle, 0xa1, new IntPtr(2), IntPtr.Zero);
            this.WndProc(ref m);
        }


        private void load_data_table()
        {
            //Получаем все данные из таблицы для соответствующего пользователя
            try
            {
                connection.Open();
                OleDbCommand command = new OleDbCommand();
                command.Connection = connection;
                command.CommandText = "select id_person,Name,Fam,mid_name,scholarship,academic_degree,academic_title from " + table_name + ";";


                OleDbDataAdapter adapter = new OleDbDataAdapter(command);
                DataTable table = new DataTable();

                Task.Run(() =>
                {
                    BeginInvoke(new MethodInvoker(delegate
                    {
                        adapter.Fill(table);
                        dataGridView.DataSource = table;

                        dataGridView.Columns[0].HeaderText = "№";
                        dataGridView.Columns[1].HeaderText = "Имя";
                        dataGridView.Columns[2].HeaderText = "Фамилия";
                        dataGridView.Columns[3].HeaderText = "Отчество";
                        dataGridView.Columns[4].HeaderText = "Учёность";
                        dataGridView.Columns[5].HeaderText = "Учёная степень";
                        dataGridView.Columns[6].HeaderText = "Учёное звание";
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


        private void clean_input()
        {
            txt_fam.Text = "";
            txt_name.Text = "";
            txt_name2.Text = "";
            comboBox_level.Text = "";
            comboBox_academic_title.Text = "";
            comboBox_academic_degree.Text = "";
        }

        private void Form_table_PPS_Load(object sender, EventArgs e)
        {
            load_data_table();
        }

    }
}
