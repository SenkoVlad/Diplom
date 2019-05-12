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
using System.IO;

namespace Science_0._01
{
    public partial class Form_login : Form
    {
        
        // строка подключения к MS Access
        // вариант 1
        public static string connectString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Application.StartupPath.ToString() + @"\science.mdb;";
        // вариант 2
        //public static string connectString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=Workers.mdb;";
        private OleDbConnection connection = new OleDbConnection();


        //Данные пользователя
        public int id_person {get; private set;}
        public string name_person { get; private set; }
        public string fam_person { get; private set; }
        public string mid_name_person { get; private set; }

        //Для чтения данных из таблицы
        OleDbDataReader reader = null;

        public Form_login()
        {
            InitializeComponent();
            connection.ConnectionString = connectString; 
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //Выполним тестовое подключение
            test_connect();
        }

        //Выполним вход и получим данные пользователя
        private void btn_Login_Click(object sender, EventArgs e)
        {
            try
            {
                id_person = 0;

                connection.Open();
                OleDbCommand command = new OleDbCommand();
                command.Connection = connection;
                command.CommandText = "select * from person where Fam='" + comboBox_fam.SelectedItem.ToString() + "' and Name='" + comboBox_name.SelectedItem.ToString() + "';";
                OleDbDataReader reader = command.ExecuteReader();

                while (reader.Read())
                {
                    id_person = Convert.ToInt32(reader[0]);
                    fam_person = reader[1].ToString();
                    name_person = reader[2].ToString();
                    mid_name_person = reader[3].ToString();
                }
              
                connection.Close();
                
                
                if (comboBox_fam.SelectedItem.ToString() == "Admin" && comboBox_name.SelectedItem.ToString() == "Admin")
                {
                    this.Hide();
                    Form_Admin form_admin = new Form_Admin();
                    form_admin.ShowDialog();
                }
                else if (id_person == 0)
                {
                    MessageBox.Show("Вас нету в БД. Обратитесь к администратору", "Упс!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                else
                {
                    this.Hide();
                    Form_selection_table form_table = new Form_selection_table(this);
                    form_table.ShowDialog();
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                connection.Close();
              //  connection.Dispose();
            }
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            System.Environment.Exit(0); // Приложение завершается и возвращает ОС указанное параметром значение
        }

        private void btn_BD_file_Click(object sender, EventArgs e)
        {
            //Выполнить тестовое подключение
            test_connect();
        }

        //Тестовое подключение к базе и выбор файла при невозможности найти БД
        private void test_connect()
        {
        connect:
            try
            {                
                connection.ConnectionString = connectString;
                connection.Open();

                OleDbCommand command = new OleDbCommand();

                command.CommandText = "select name,fam from person;";
                command.Connection = connection;
                reader = command.ExecuteReader();

                comboBox_fam.Items.Clear();
                comboBox_name.Items.Clear();

                while (reader.Read())
                {
                    if (!comboBox_name.Items.Contains(reader[0].ToString()))
                        comboBox_name.Items.Add(reader[0].ToString());
                    if (!comboBox_fam.Items.Contains(reader[1].ToString()))
                        comboBox_fam.Items.Add(reader[1].ToString());
                }
                comboBox_name.Items.Add("Admin");
                comboBox_fam.Items.Add("Admin");

                reader.Close();

                img_yes_conn.Visible = true;
                btn_BD_file.Visible = false;
                img_no_conn.Visible = false;
                connection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                MessageBox.Show("Попробуйсте выбрать БД", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                var filePath = string.Empty;

                using (OpenFileDialog openFileDialog = new OpenFileDialog())
                {
                    openFileDialog.RestoreDirectory = true;

                    if (openFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        //Get the path of specified file
                        connectString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + openFileDialog.FileName + ";";
                        //connectString =  connectString.Replace("\\", @"\");
                        goto connect;
                    }
                    else
                    {
                        MessageBox.Show("Не найдена БД", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
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

        private void Form_login_MouseDown(object sender, MouseEventArgs e)
        {
            base.Capture = false;
            Message m = Message.Create(base.Handle, 0xa1, new IntPtr(2), IntPtr.Zero);
            this.WndProc(ref m);
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }
    }
}
