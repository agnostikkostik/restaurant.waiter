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

namespace WindowsFormsApplication1
{
    public partial class MsgBoxYesNo : Form
    {
        OleDbCommand komand;
        OleDbDataReader post;
        OleDbConnection podkl = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + @"\\PC-SPB\Users\RESTORAN\baza.mdb");
        /// <summary>
        /// 
        /// </summary>
        /// <param name="temp1">Текст</param>
        /// <param name="temp2">1 - ОК, 2 - Да/Нет, 3 - Удаление, 4 - Причина</param>
        public MsgBoxYesNo(string temp1, int temp2)
        {
            InitializeComponent();
            textBox1.Text = temp1;
            if (temp2 == 1)
            {
                hide_all();
                ok.Visible = true;
                ok.Text = "ОК";
            }
            if (temp2 == 2)
            {
                hide_all();
                yes.Visible = true;
                no.Visible = true;
                yes.Text = "Да";
                no.Text = "Нет";
            }
            if (temp2 == 3)
            {
                hide_all();
                prichina_1.Visible = true;
                prichina_2.Visible = true;
                prichina_3.Visible = true;
                prichina_4.Visible = true;
                prichina_1.Text = "Отказ гостя";
                prichina_2.Text = "Ошибка официанта";
                prichina_3.Text = "Стоп лист";
                prichina_4.Text = "Не приготовили";
            }
            if (temp2 == 4)
            {
                hide_all();
                yes.Visible = true;
                ok.Visible = true;
                no.Visible = true;
                yes.Text = "Удалить";
                ok.Text = "Переместить";
                no.Text = "Отмена";
            }
        }

        private void hide_all()
        {
            ok.Visible = false;
            yes.Visible = false;
            no.Visible = false;
            prichina_1.Visible = false;
            prichina_2.Visible = false;
            prichina_3.Visible = false;
            prichina_4.Visible = false;
        }

        private void MsgBoxYesNo_Load(object sender, EventArgs e)
        {
            Button[] all_buttons = new Button[] { ok, yes, no,
                prichina_1, prichina_2, prichina_3, prichina_4
            };
            for (int i = 0; i < all_buttons.Length; i++)
            {
                all_buttons[i].BackColor = Color.FromArgb(255, 215, 228, 242);
                all_buttons[i].FlatStyle = FlatStyle.Flat;
            }
            this.BackColor = Color.FromArgb(255, 120, 215, 124);
            textBox1.BackColor = Color.FromArgb(255, 120, 215, 124);
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            podkl.Open();
            bool f = false;
            komand = new OleDbCommand("Select [id сотрудника], [id должности] From [Сотрудники] WHERE [EAN13]='" + textBox2.Text + "'", podkl);
            post = komand.ExecuteReader();
            if (post.Read() != false)
                f = true;
            
            if (f)
            {
                if ((post.GetValue(1).ToString() != "ОФ") && (post.GetValue(1).ToString() != "СО"))
                {
                    (Owner as Form1).razreshenie = post.GetValue(0).ToString();
                    button1.PerformClick();
                }
            }
            podkl.Close();
        }
    }
}
