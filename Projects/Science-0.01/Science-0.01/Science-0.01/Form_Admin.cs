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
    public partial class Form_Admin : Form
    {
        public Form_Admin()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

        }



        private void Form_Admin_Load(object sender, EventArgs e)
        {

        }

        private void btn_table_Click(object sender, EventArgs e)
        {
            this.Hide();
            Form_table_PPS form_PPS = new Form_table_PPS();
            form_PPS.ShowDialog();
        }

        private void Form_Admin_MouseDown(object sender, MouseEventArgs e)
        {
            base.Capture = false;
            Message m = Message.Create(base.Handle, 0xa1, new IntPtr(2), IntPtr.Zero);
            this.WndProc(ref m);
        }

        private void btn_exit_2_Click(object sender, EventArgs e)
        {
            System.Environment.Exit(0); // Приложение завершается и возвращает ОС указанное параметром значение

        }

        private void Form_Admin_FormClosed(object sender, FormClosedEventArgs e)
        {
            System.Environment.Exit(0); // Приложение завершается и возвращает ОС указанное параметром значение
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

        private void btn_back_Click(object sender, EventArgs e)
        {
            this.Hide();
            Form_login form_log = new Form_login();
            form_log.ShowDialog();
        }

        private void btn_report_Click(object sender, EventArgs e)
        {
            this.Hide();
            Form_reports form_report = new Form_reports();
            form_report.ShowDialog();
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            this.Hide();
            Form_ListConf form_conf = new Form_ListConf();
            form_conf.ShowDialog();
        }
    }
}
