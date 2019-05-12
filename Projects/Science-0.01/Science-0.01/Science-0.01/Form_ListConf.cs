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
    public partial class Form_ListConf : Form
    {
        private OleDbConnection connection = new OleDbConnection();


        //флаг для определения хочет ли пользователь обновить строку
        private bool flag_update = false;
        private bool flag_update2 = false;


        //для отслеживание выбранной статьи
        private string id_record_conf = "0";
        private string id_record_article = "0";

        public Form_ListConf()
        {
            InitializeComponent();
            connection.ConnectionString = Form_login.connectString;
        }

        private void btn_back_Click(object sender, EventArgs e)
        {
            this.Hide();
            Form_Admin form_admin = new Form_Admin();
            form_admin.ShowDialog();
        }

        private void Form_ListConf_MouseDown(object sender, MouseEventArgs e)
        {
            base.Capture = false;
            Message m = Message.Create(base.Handle, 0xa1, new IntPtr(2), IntPtr.Zero);
            this.WndProc(ref m);
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

        private void btn_update_Click(object sender, EventArgs e)
        {
            panel.Enabled = true;
            flag_update = true;
        }

        private void Form_ListConf_Load(object sender, EventArgs e)
        {
            load_data_table_2();
            load_data_table();
        }

        private void btn_update_2_Click(object sender, EventArgs e)
        {
            panel1.Enabled = true;
            flag_update2 = true;
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


                    command.CommandText = "delete from conference_list where id_record_conf=" + int.Parse(id_record_conf) + ";";



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
                        command.CommandText = "update conference_list set name_conference='" + txt_conf.Text + "'," +  "place_conference='" + txt_place.Text + "' where id_record_conf="+int.Parse(id_record_conf)+";";
                    }
                    else
                    {
                        command.CommandText = "insert into conference_list (name_conference,place_conference) values('" + txt_conf.Text + "','" + txt_place.Text+"');";
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

        private void load_data_table_2()
        {

            //Получаем все данные из таблицы для соответствующего пользователя
            try
            {
                connection.Open();
                OleDbCommand command = new OleDbCommand();
                command.Connection = connection;

                command.CommandText = "select id_article_record,title_edition from articles_list;";

                OleDbDataAdapter adapter1 = new OleDbDataAdapter(command);
                DataTable table1 = new DataTable();

                Task.Run(() =>
                {
                    BeginInvoke(new MethodInvoker(delegate
                    {
                        adapter1.Fill(table1);
                        dataGridView1.DataSource = table1;

                        dataGridView1.Columns[0].HeaderText = "№";
                        dataGridView1.Columns[1].HeaderText = "Название издания";
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

        private void load_data_table()
        {
            //Получаем все данные из таблицы для соответствующего пользователя
            try
            {
                connection.Open();
                OleDbCommand command = new OleDbCommand();
                command.Connection = connection;

                command.CommandText = "select id_record_conf,name_conference,place_conference from conference_list;";


                OleDbDataAdapter adapter = new OleDbDataAdapter(command);
                DataTable table = new DataTable();

                Task.Run(() =>
                {
                    BeginInvoke(new MethodInvoker(delegate
                    {
                        adapter.Fill(table);
                        dataGridView.DataSource = table;

                        dataGridView.Columns[0].HeaderText = "№";
                        dataGridView.Columns[1].HeaderText = "Название конференции";
                        dataGridView.Columns[2].HeaderText = "Место проведения";
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
            txt_place.Text = "";
            txt_conf.Text = "";
        }
        private void clean_input_2()
        {
            txt_edition.Text = "";
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
                id_record_conf = row.Cells[0].Value.ToString();
                txt_conf.Text = row.Cells[1].Value.ToString();
                txt_place.Text = row.Cells[2].Value.ToString();
            }
        }

        private void btn_new_1_Click(object sender, EventArgs e)
        {
            panel1.Enabled = true;
            flag_update2 = false;
            clean_input_2();
        }

        private async void btn_save_2_Click(object sender, EventArgs e)
        {
            if (panel1.Enabled)
            {
                //Вносим новую строку
                try
                {
                    connection.Open();
                    OleDbCommand command = new OleDbCommand();
                    command.Connection = connection;

                    if (flag_update2)
                    {
                        command.CommandText = "update articles_list set title_edition='" + txt_edition.Text + "' where id_article_record=" + int.Parse(id_record_article) + ";";
                    }
                    else
                    {
                        command.CommandText = "insert into articles_list (title_edition) values('" + txt_edition.Text + "');";
                    }


                    await Task.Run(() => command.ExecuteNonQuery());

                    if (flag_update2)
                        show_info("Данные обновлены", Color.Gold, 2000);
                    else
                        show_info("Данные внесены", Color.GreenYellow, 2000);

                    panel1.Enabled = false;
                    flag_update2 = false;
                    connection.Close();

                    //Обновить таблицу
                    load_data_table_2();
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

        private async void btn_delete_2_Click(object sender, EventArgs e)
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


                    command.CommandText = "delete from articles_list where id_article_record=" + int.Parse(id_record_article) + ";";



                    await Task.Run(() => command.ExecuteNonQuery());

                    show_info("Запись удалена", Color.Red, 2000);

                    panel.Enabled = false;
                    connection.Close();
                    //Обновить таблицу
                    load_data_table_2();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    connection.Close();
                }
            }
        }

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            panel1.Enabled = false;

            DataGridViewCell cell = null;
            foreach (DataGridViewCell selectedCell in dataGridView1.SelectedCells)
            {
                cell = selectedCell;
                break;
            }
            if (cell != null)
            {
                DataGridViewRow row = cell.OwningRow;
                id_record_article = row.Cells[0].Value.ToString();
                txt_edition.Text = row.Cells[1].Value.ToString();
            }
        }
    }
}
