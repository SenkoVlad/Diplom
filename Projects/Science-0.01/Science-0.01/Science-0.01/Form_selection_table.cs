using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Science_0._01
{
    public partial class Form_selection_table : Form
    {


        //Данные пользователя
        Form_login login_form;

        //Имя таблица, для внесения данных
        private string table_name = "";



        public Form_selection_table()
        {
            InitializeComponent();
        }

        //Получим данные пользователя
        public Form_selection_table(Form_login form_login)
        {
            InitializeComponent();
            login_form = form_login;
            label_user_name.Text = login_form.fam_person + " " + login_form.name_person + " " + login_form.mid_name_person;
        }

        private void Form_selection_table_FormClosed(object sender, FormClosedEventArgs e)
        {
            System.Environment.Exit(0); // Приложение завершается и возвращает ОС указанное параметром значение 
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Hide();
            Form_login form_login = new Form_login();
            form_login.ShowDialog();
        }

        //Получаем название выбранной таблицы и запоминаем
        private void button1_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < groupBox_table.Controls.Count; i++)
            {
                if (((RadioButton)groupBox_table.Controls[i]).Checked == true) 
                {
                    table_name = ((RadioButton)groupBox_table.Controls[i]).Name;

                    switch (table_name)
                    {
                        case "conference":
                            this.Hide();
                            Form_conf form_conf = new Form_conf(login_form, table_name);
                            form_conf.ShowDialog();
                            break;
                        case "articles":
                            this.Hide();
                            Form_article form_article = new Form_article(login_form, table_name);
                            form_article.ShowDialog();
                            break;
                        case "textbooks":
                            this.Hide();
                            Form_textbooks form_textbook= new Form_textbooks(login_form, table_name);
                            form_textbook.ShowDialog();
                            break;
                        case "military_training":
                            this.Hide();
                            Form_mill_training form_training = new Form_mill_training(login_form, table_name);
                            form_training.ShowDialog();
                            break;
                        case "reasearch_work":
                            this.Hide();
                            Form_research_work form_reasearch = new Form_research_work(login_form, table_name);
                            form_reasearch.ShowDialog();
                            break;
                        case "inventions":
                            this.Hide();
                            Form_inventions form_invention = new Form_inventions(login_form, table_name);
                            form_invention.ShowDialog();
                            break;
                        case "seminar":
                            this.Hide();
                            Form_seminar form_seminar = new Form_seminar(login_form, table_name);
                            form_seminar.ShowDialog();
                            break;
                        case "reviews":
                            this.Hide();
                            Form_review form_review = new Form_review(login_form, table_name);
                            form_review.ShowDialog();
                            break;
                        case "conf_kursant":
                            this.Hide();
                            Form_conf_kursant form_conf_kursant = new Form_conf_kursant(login_form, table_name);
                            form_conf_kursant.ShowDialog();
                            break;
                        case "scientific_leaders":
                            this.Hide();
                            Form_scientific_leaders form_scientific_leaders = new Form_scientific_leaders(login_form, table_name);
                            form_scientific_leaders.ShowDialog();
                            break;
                        case "exhibitions":
                            this.Hide();
                            Form_science_work form_science_work = new Form_science_work(login_form, table_name);
                            form_science_work.ShowDialog();
                            break;
                        default:
                            MessageBox.Show("Выберите что-нибудь", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            break;
                    }

                    
                }
            }
        }

        private void Form_selection_table_Load(object sender, EventArgs e)
        {

        }

        private void conference_CheckedChanged(object sender, EventArgs e)
        {

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

        private void Form_selection_table_MouseDown(object sender, MouseEventArgs e)
        {
            base.Capture = false;
            Message m = Message.Create(base.Handle, 0xa1, new IntPtr(2), IntPtr.Zero);
            this.WndProc(ref m);
        }

        private void PaintBorderlessGroupBox(object sender, PaintEventArgs p)
        {
        }

        private void conf_kursant_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void seminar_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void scientific_leaders_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void reviews_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void exhibitions_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void textbooks_CheckedChanged(object sender, EventArgs e)
        {

        }

    }
}
