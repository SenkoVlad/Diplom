namespace Science_0._01
{
    partial class Form_login
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form_login));
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.btn_exit = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.btn_Login = new System.Windows.Forms.Button();
            this.label_conn_status = new System.Windows.Forms.Label();
            this.btn_BD_file = new System.Windows.Forms.Button();
            this.btn_minimize = new System.Windows.Forms.Button();
            this.btn_restore = new System.Windows.Forms.Button();
            this.btn_exit_2 = new System.Windows.Forms.Button();
            this.img_no_conn = new System.Windows.Forms.PictureBox();
            this.img_yes_conn = new System.Windows.Forms.PictureBox();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.comboBox_name = new System.Windows.Forms.ComboBox();
            this.comboBox_fam = new System.Windows.Forms.ComboBox();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.img_no_conn)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.img_yes_conn)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.comboBox_fam);
            this.groupBox1.Controls.Add(this.comboBox_name);
            this.groupBox1.Controls.Add(this.btn_exit);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.btn_Login);
            this.groupBox1.Font = new System.Drawing.Font("Times New Roman", 18F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.groupBox1.ForeColor = System.Drawing.Color.PaleTurquoise;
            this.groupBox1.Location = new System.Drawing.Point(275, 39);
            this.groupBox1.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Padding = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.groupBox1.Size = new System.Drawing.Size(670, 458);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Представьтесь";
            // 
            // btn_exit
            // 
            this.btn_exit.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btn_exit.Font = new System.Drawing.Font("Times New Roman", 18F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.btn_exit.ForeColor = System.Drawing.Color.PaleTurquoise;
            this.btn_exit.Image = global::Science_0._01.Properties.Resources.logout_icon__1_;
            this.btn_exit.ImageAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.btn_exit.Location = new System.Drawing.Point(2, 385);
            this.btn_exit.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.btn_exit.Name = "btn_exit";
            this.btn_exit.Size = new System.Drawing.Size(119, 71);
            this.btn_exit.TabIndex = 6;
            this.btn_exit.Text = "Выход";
            this.btn_exit.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            this.btn_exit.UseVisualStyleBackColor = true;
            this.btn_exit.Click += new System.EventHandler(this.btn_exit_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Times New Roman", 18F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label2.ForeColor = System.Drawing.Color.PaleTurquoise;
            this.label2.Location = new System.Drawing.Point(10, 207);
            this.label2.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(115, 26);
            this.label2.TabIndex = 5;
            this.label2.Text = "Фамилия";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Times New Roman", 18F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label1.ForeColor = System.Drawing.Color.PaleTurquoise;
            this.label1.Location = new System.Drawing.Point(10, 124);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(60, 26);
            this.label1.TabIndex = 4;
            this.label1.Text = "Имя";
            this.label1.Click += new System.EventHandler(this.label1_Click);
            // 
            // btn_Login
            // 
            this.btn_Login.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btn_Login.Font = new System.Drawing.Font("Times New Roman", 18F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.btn_Login.ForeColor = System.Drawing.Color.PaleTurquoise;
            this.btn_Login.Image = global::Science_0._01.Properties.Resources.Login_icon;
            this.btn_Login.ImageAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.btn_Login.Location = new System.Drawing.Point(546, 385);
            this.btn_Login.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.btn_Login.Name = "btn_Login";
            this.btn_Login.Size = new System.Drawing.Size(122, 71);
            this.btn_Login.TabIndex = 3;
            this.btn_Login.Text = "Вход";
            this.btn_Login.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            this.btn_Login.UseVisualStyleBackColor = true;
            this.btn_Login.Click += new System.EventHandler(this.btn_Login_Click);
            // 
            // label_conn_status
            // 
            this.label_conn_status.AutoSize = true;
            this.label_conn_status.Font = new System.Drawing.Font("Times New Roman", 18F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label_conn_status.ForeColor = System.Drawing.Color.PaleTurquoise;
            this.label_conn_status.Location = new System.Drawing.Point(4, 314);
            this.label_conn_status.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label_conn_status.MaximumSize = new System.Drawing.Size(274, 0);
            this.label_conn_status.Name = "label_conn_status";
            this.label_conn_status.Size = new System.Drawing.Size(257, 52);
            this.label_conn_status.TabIndex = 3;
            this.label_conn_status.Text = "Статус подключения к БД";
            this.label_conn_status.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // btn_BD_file
            // 
            this.btn_BD_file.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btn_BD_file.Font = new System.Drawing.Font("Times New Roman", 18F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.btn_BD_file.ForeColor = System.Drawing.Color.PaleTurquoise;
            this.btn_BD_file.Location = new System.Drawing.Point(0, 370);
            this.btn_BD_file.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.btn_BD_file.Name = "btn_BD_file";
            this.btn_BD_file.Size = new System.Drawing.Size(116, 89);
            this.btn_BD_file.TabIndex = 5;
            this.btn_BD_file.Text = "Выбрать БД";
            this.btn_BD_file.UseVisualStyleBackColor = true;
            this.btn_BD_file.Click += new System.EventHandler(this.btn_BD_file_Click);
            // 
            // btn_minimize
            // 
            this.btn_minimize.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btn_minimize.Image = global::Science_0._01.Properties.Resources.minimize_icon;
            this.btn_minimize.Location = new System.Drawing.Point(830, 0);
            this.btn_minimize.Name = "btn_minimize";
            this.btn_minimize.Size = new System.Drawing.Size(38, 37);
            this.btn_minimize.TabIndex = 9;
            this.btn_minimize.UseVisualStyleBackColor = true;
            this.btn_minimize.Click += new System.EventHandler(this.btn_minimize_Click);
            // 
            // btn_restore
            // 
            this.btn_restore.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btn_restore.Image = global::Science_0._01.Properties.Resources.restore_icon;
            this.btn_restore.Location = new System.Drawing.Point(867, 0);
            this.btn_restore.Name = "btn_restore";
            this.btn_restore.Size = new System.Drawing.Size(40, 37);
            this.btn_restore.TabIndex = 8;
            this.btn_restore.UseVisualStyleBackColor = true;
            this.btn_restore.Click += new System.EventHandler(this.btn_restore_Click);
            // 
            // btn_exit_2
            // 
            this.btn_exit_2.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btn_exit_2.Image = global::Science_0._01.Properties.Resources.Actions_window_close_icon;
            this.btn_exit_2.Location = new System.Drawing.Point(906, 0);
            this.btn_exit_2.Name = "btn_exit_2";
            this.btn_exit_2.Size = new System.Drawing.Size(40, 37);
            this.btn_exit_2.TabIndex = 7;
            this.btn_exit_2.UseVisualStyleBackColor = true;
            this.btn_exit_2.Click += new System.EventHandler(this.btn_exit_2_Click);
            // 
            // img_no_conn
            // 
            this.img_no_conn.Image = global::Science_0._01.Properties.Resources.Close_2_icon;
            this.img_no_conn.Location = new System.Drawing.Point(117, 370);
            this.img_no_conn.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.img_no_conn.Name = "img_no_conn";
            this.img_no_conn.Size = new System.Drawing.Size(115, 89);
            this.img_no_conn.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.img_no_conn.TabIndex = 4;
            this.img_no_conn.TabStop = false;
            // 
            // img_yes_conn
            // 
            this.img_yes_conn.Image = global::Science_0._01.Properties.Resources.Ok_icon;
            this.img_yes_conn.Location = new System.Drawing.Point(116, 370);
            this.img_yes_conn.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.img_yes_conn.Name = "img_yes_conn";
            this.img_yes_conn.Size = new System.Drawing.Size(115, 89);
            this.img_yes_conn.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.img_yes_conn.TabIndex = 2;
            this.img_yes_conn.TabStop = false;
            this.img_yes_conn.Visible = false;
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = global::Science_0._01.Properties.Resources.Teachers_icon;
            this.pictureBox1.Location = new System.Drawing.Point(0, 52);
            this.pictureBox1.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(239, 210);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBox1.TabIndex = 1;
            this.pictureBox1.TabStop = false;
            // 
            // comboBox_name
            // 
            this.comboBox_name.BackColor = System.Drawing.SystemColors.InactiveCaptionText;
            this.comboBox_name.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.comboBox_name.Font = new System.Drawing.Font("Times New Roman", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.comboBox_name.ForeColor = System.Drawing.Color.PaleTurquoise;
            this.comboBox_name.FormattingEnabled = true;
            this.comboBox_name.Items.AddRange(new object[] {
            "Тема",
            "Руководитель",
            "Руководимый"});
            this.comboBox_name.Location = new System.Drawing.Point(167, 123);
            this.comboBox_name.Name = "comboBox_name";
            this.comboBox_name.Size = new System.Drawing.Size(426, 27);
            this.comboBox_name.TabIndex = 18;
            // 
            // comboBox_fam
            // 
            this.comboBox_fam.BackColor = System.Drawing.SystemColors.InactiveCaptionText;
            this.comboBox_fam.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.comboBox_fam.Font = new System.Drawing.Font("Times New Roman", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.comboBox_fam.ForeColor = System.Drawing.Color.PaleTurquoise;
            this.comboBox_fam.FormattingEnabled = true;
            this.comboBox_fam.Items.AddRange(new object[] {
            "Тема",
            "Руководитель",
            "Руководимый"});
            this.comboBox_fam.Location = new System.Drawing.Point(167, 207);
            this.comboBox_fam.Name = "comboBox_fam";
            this.comboBox_fam.Size = new System.Drawing.Size(426, 27);
            this.comboBox_fam.TabIndex = 19;
            // 
            // Form_login
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(10F, 23F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(41)))), ((int)(((byte)(44)))), ((int)(((byte)(51)))));
            this.ClientSize = new System.Drawing.Size(946, 497);
            this.Controls.Add(this.btn_minimize);
            this.Controls.Add(this.btn_restore);
            this.Controls.Add(this.btn_exit_2);
            this.Controls.Add(this.btn_BD_file);
            this.Controls.Add(this.img_no_conn);
            this.Controls.Add(this.label_conn_status);
            this.Controls.Add(this.img_yes_conn);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.groupBox1);
            this.Font = new System.Drawing.Font("Comic Sans MS", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.ForeColor = System.Drawing.Color.PaleTurquoise;
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.Name = "Form_login";
            this.Text = "Вход";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.Form1_FormClosed);
            this.Load += new System.EventHandler(this.Form1_Load);
            this.MouseDown += new System.Windows.Forms.MouseEventHandler(this.Form_login_MouseDown);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.img_no_conn)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.img_yes_conn)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btn_Login;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.PictureBox img_yes_conn;
        private System.Windows.Forms.Label label_conn_status;
        private System.Windows.Forms.PictureBox img_no_conn;
        private System.Windows.Forms.Button btn_BD_file;
        private System.Windows.Forms.Button btn_exit;
        private System.Windows.Forms.Button btn_exit_2;
        private System.Windows.Forms.Button btn_restore;
        private System.Windows.Forms.Button btn_minimize;
        private System.Windows.Forms.ComboBox comboBox_fam;
        private System.Windows.Forms.ComboBox comboBox_name;
    }
}

