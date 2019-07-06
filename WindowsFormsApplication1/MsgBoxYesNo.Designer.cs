namespace WindowsFormsApplication1
{
    partial class MsgBoxYesNo
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
            this.yes = new System.Windows.Forms.Button();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.no = new System.Windows.Forms.Button();
            this.ok = new System.Windows.Forms.Button();
            this.prichina_1 = new System.Windows.Forms.Button();
            this.prichina_2 = new System.Windows.Forms.Button();
            this.prichina_3 = new System.Windows.Forms.Button();
            this.prichina_4 = new System.Windows.Forms.Button();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.button1 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // yes
            // 
            this.yes.DialogResult = System.Windows.Forms.DialogResult.Yes;
            this.yes.Location = new System.Drawing.Point(12, 41);
            this.yes.Name = "yes";
            this.yes.Size = new System.Drawing.Size(104, 57);
            this.yes.TabIndex = 1;
            this.yes.TabStop = false;
            this.yes.Text = "Да";
            this.yes.UseVisualStyleBackColor = true;
            // 
            // textBox1
            // 
            this.textBox1.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.textBox1.Location = new System.Drawing.Point(12, 12);
            this.textBox1.Multiline = true;
            this.textBox1.Name = "textBox1";
            this.textBox1.ReadOnly = true;
            this.textBox1.Size = new System.Drawing.Size(448, 23);
            this.textBox1.TabIndex = 4;
            this.textBox1.TabStop = false;
            this.textBox1.Text = "Недостаточно прав. Необходимо менджерское подтверждение.";
            this.textBox1.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // no
            // 
            this.no.DialogResult = System.Windows.Forms.DialogResult.No;
            this.no.Location = new System.Drawing.Point(356, 41);
            this.no.Name = "no";
            this.no.Size = new System.Drawing.Size(104, 57);
            this.no.TabIndex = 2;
            this.no.TabStop = false;
            this.no.Text = "Нет";
            this.no.UseVisualStyleBackColor = true;
            // 
            // ok
            // 
            this.ok.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.ok.Location = new System.Drawing.Point(184, 41);
            this.ok.Name = "ok";
            this.ok.Size = new System.Drawing.Size(104, 56);
            this.ok.TabIndex = 3;
            this.ok.TabStop = false;
            this.ok.Text = "ОК";
            this.ok.UseVisualStyleBackColor = true;
            // 
            // prichina_1
            // 
            this.prichina_1.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.prichina_1.Location = new System.Drawing.Point(19, 41);
            this.prichina_1.Name = "prichina_1";
            this.prichina_1.Size = new System.Drawing.Size(104, 56);
            this.prichina_1.TabIndex = 5;
            this.prichina_1.TabStop = false;
            this.prichina_1.Text = "Отказ гостя";
            this.prichina_1.UseVisualStyleBackColor = true;
            // 
            // prichina_2
            // 
            this.prichina_2.DialogResult = System.Windows.Forms.DialogResult.Abort;
            this.prichina_2.Location = new System.Drawing.Point(129, 41);
            this.prichina_2.Name = "prichina_2";
            this.prichina_2.Size = new System.Drawing.Size(104, 57);
            this.prichina_2.TabIndex = 6;
            this.prichina_2.TabStop = false;
            this.prichina_2.Text = "Ошибка официанта";
            this.prichina_2.UseVisualStyleBackColor = true;
            // 
            // prichina_3
            // 
            this.prichina_3.DialogResult = System.Windows.Forms.DialogResult.Retry;
            this.prichina_3.Location = new System.Drawing.Point(239, 41);
            this.prichina_3.Name = "prichina_3";
            this.prichina_3.Size = new System.Drawing.Size(104, 57);
            this.prichina_3.TabIndex = 8;
            this.prichina_3.TabStop = false;
            this.prichina_3.Text = "Стоп лист";
            this.prichina_3.UseVisualStyleBackColor = true;
            // 
            // prichina_4
            // 
            this.prichina_4.DialogResult = System.Windows.Forms.DialogResult.Ignore;
            this.prichina_4.Location = new System.Drawing.Point(349, 42);
            this.prichina_4.Name = "prichina_4";
            this.prichina_4.Size = new System.Drawing.Size(104, 56);
            this.prichina_4.TabIndex = 7;
            this.prichina_4.TabStop = false;
            this.prichina_4.Text = "Не приготовили";
            this.prichina_4.UseVisualStyleBackColor = true;
            // 
            // textBox2
            // 
            this.textBox2.Location = new System.Drawing.Point(16, 140);
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(363, 20);
            this.textBox2.TabIndex = 9;
            this.textBox2.TextChanged += new System.EventHandler(this.textBox2_TextChanged);
            // 
            // button1
            // 
            this.button1.DialogResult = System.Windows.Forms.DialogResult.Yes;
            this.button1.Location = new System.Drawing.Point(385, 140);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 10;
            this.button1.TabStop = false;
            this.button1.Text = "button1";
            this.button1.UseVisualStyleBackColor = true;
            // 
            // MsgBoxYesNo
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(472, 109);
            this.ControlBox = false;
            this.Controls.Add(this.button1);
            this.Controls.Add(this.textBox2);
            this.Controls.Add(this.prichina_3);
            this.Controls.Add(this.prichina_4);
            this.Controls.Add(this.prichina_2);
            this.Controls.Add(this.prichina_1);
            this.Controls.Add(this.ok);
            this.Controls.Add(this.no);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.yes);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.SizableToolWindow;
            this.MaximumSize = new System.Drawing.Size(488, 143);
            this.MinimumSize = new System.Drawing.Size(488, 143);
            this.Name = "MsgBoxYesNo";
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Сообщение";
            this.Load += new System.EventHandler(this.MsgBoxYesNo_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Button yes;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Button no;
        private System.Windows.Forms.Button ok;
        private System.Windows.Forms.Button prichina_1;
        private System.Windows.Forms.Button prichina_2;
        private System.Windows.Forms.Button prichina_3;
        private System.Windows.Forms.Button prichina_4;
        private System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.Button button1;
    }
}