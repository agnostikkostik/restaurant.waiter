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
using System.Net;

namespace WindowsFormsApplication1
{
    public partial class auth : Form
    {
        OleDbCommand komand;
        OleDbDataReader post;
        OleDbConnection podkl = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + @"\\PC-SPB\Users\RESTORAN\baza.mdb");

        int focus;
        TextBox[] temp;
        public string kod;

        public auth()
        {
            InitializeComponent();

            #region РАЗМЕРЫ
            textBox1.BackColor = Color.FromArgb(255, 120, 215, 124);
            temp = new TextBox[] { login, password, EAN };

            int h = Screen.PrimaryScreen.Bounds.Height;
            int w = Screen.PrimaryScreen.Bounds.Width;

            this.MaximumSize = new Size(w, h);
            this.MinimumSize = new Size(w, h);
            this.Size = new Size(w, h);

            label1.Location = new Point((login.Location.X - 32), (login.Location.Y + 3));
            label2.Location = new Point((password.Location.X - 51), (password.Location.Y + 3));
            label3.Location = new Point((EAN.Location.X - 47), (EAN.Location.Y + 3));

            textBox1.Location = new Point((login.Location.X - 19), (login.Location.Y - 63));
            #endregion
        }

        private void login_button_Click(object sender, EventArgs e)
        {
            podkl.Open();
            if (EAN.Text == "")
            {
                komand = new OleDbCommand("Select [id сотрудника], [Фамилия], [Имя], [id должности] From [Сотрудники] WHERE [Логин]='" + login.Text + "' AND [Пароль]='" + password.Text + "'", podkl);
                post = komand.ExecuteReader();
                if (post.Read() == false)
                {
                    MessageBox.Show("Сотрудника с такими данными нет!");
                    podkl.Close();
                    return;
                }

                kod = post.GetValue(0).ToString();
                string FI = post.GetValue(1).ToString() + " " + post.GetValue(2).ToString();
                string id = post.GetValue(3).ToString();

                if (textBox1.Visible == true)
                {
                    if ((id == "ОФ") || (id == "СО"))
                    {
                        new MsgBoxYesNo("НЕДОСТАТОЧНО ПРАВ.", 1).ShowDialog();
                        podkl.Close();
                        return;
                    }
                    else
                    {
                        komand = new OleDbCommand("SELECT Count(*) From [Заказы]", podkl);
                        post = komand.ExecuteReader();
                        post.Read();
                        if (post.GetValue(0).ToString() == "0")
                        {
                            komand = new OleDbCommand("INSERT INTO [Заказы]([Код заказа], [Код сотрудника], [Номер стола], [Когда открыт]) VALUES (0, '" + kod + "', 999999, '" + DateTime.Now.ToString() + "')", podkl);
                            post = komand.ExecuteReader();
                        }
                        textBox1.Visible = false;
                    }
                }

                select_stol frm = new select_stol(kod, id, FI);
                frm.Owner = this;
                frm.Show();
                login.Text = "";
                password.Text = "";
                EAN.Text = "";
                login.Focus();
                podkl.Close();
                this.Hide();
            }
            else
            {
                komand = new OleDbCommand("Select [id сотрудника], [Фамилия], [Имя], [id должности] From [Сотрудники] WHERE [EAN13]='" + EAN.Text + "'", podkl);
                post = komand.ExecuteReader();
                if (post.Read() == false)
                {
                    MessageBox.Show("Сотрудника с такими данными нет!");
                    podkl.Close();
                    return;
                }

                kod = post.GetValue(0).ToString();
                string FI = post.GetValue(1).ToString() + " " + post.GetValue(2).ToString();
                string id = post.GetValue(3).ToString();

                if (textBox1.Visible == true)
                {
                    if ((id == "ОФ") || (id == "СО"))
                    {
                        new MsgBoxYesNo("НЕДОСТАТОЧНО ПРАВ.", 1).ShowDialog();
                        podkl.Close();
                        return;
                    }
                    else
                    {
                        komand = new OleDbCommand("SELECT Count(*) From [Заказы]", podkl);
                        post = komand.ExecuteReader();
                        post.Read();
                        if (post.GetValue(0).ToString() == "0")
                        {
                            komand = new OleDbCommand("INSERT INTO [Заказы]([Код заказа], [Код сотрудника], [Номер стола], [Когда открыт]) VALUES (0, '" + kod + "', 999999, '" + DateTime.Now.ToString() + "')", podkl);
                            post = komand.ExecuteReader();
                        }
                        textBox1.Visible = false;
                    }
                }

                select_stol frm = new select_stol(kod, id, FI);
                frm.Owner = this;
                frm.Show();
                login.Text = "";
                password.Text = "";
                login.Focus();
                podkl.Close();
                this.Hide();
            }
        }

        private void login_Click(object sender, EventArgs e)
        {
            focus = 0;
        }

        private void password_Click(object sender, EventArgs e)
        {
            focus = 1;
        }

        private void EAN_Click(object sender, EventArgs e)
        {
            focus = 2;
        }

        private void kalc_Click(object sender, EventArgs e)
        {
            switch ((sender as Button).Text)
            {
                case "DEL":
                    temp[focus].Text = "";
                    break;
                case "<-":
                    if (temp[focus].Text != "")
                        temp[focus].Text = temp[focus].Text.Substring(0, temp[focus].Text.Length - 1);
                    break;
                default:
                    temp[focus].Text += (sender as Button).Text;
                    break;
            }
            temp[focus].Focus();
        }

        private void auth_Load(object sender, EventArgs e)
        {
            #region ЦВЕТА
            button1.BackColor = Color.AliceBlue;
            button1.BackColor = Color.AliceBlue;
            button2.BackColor = Color.AntiqueWhite;
            button3.BackColor = Color.Aqua;
            button4.BackColor = Color.Aquamarine;
            button5.BackColor = Color.Azure;
            button6.BackColor = Color.Beige;
            button7.BackColor = Color.Bisque;
            button8.BackColor = Color.Black;
            button9.BackColor = Color.BlanchedAlmond;
            button10.BackColor = Color.Blue;
            button11.BackColor = Color.BlueViolet;
            button12.BackColor = Color.Brown;
            button13.BackColor = Color.BurlyWood;
            button14.BackColor = Color.CadetBlue;
            button15.BackColor = Color.Chartreuse;
            button16.BackColor = Color.Chocolate;
            button17.BackColor = Color.Coral;
            button18.BackColor = Color.CornflowerBlue;
            button19.BackColor = Color.Cornsilk;
            button20.BackColor = Color.Crimson;
            button21.BackColor = Color.Cyan;
            button22.BackColor = Color.DarkBlue;
            button23.BackColor = Color.DarkCyan;
            button24.BackColor = Color.DarkGoldenrod;
            button25.BackColor = Color.DarkGray;
            button26.BackColor = Color.DarkGreen;
            button27.BackColor = Color.DarkKhaki;
            button28.BackColor = Color.DarkMagenta;
            button29.BackColor = Color.DarkOliveGreen;
            button30.BackColor = Color.DarkOrange;
            button31.BackColor = Color.DarkOrchid;
            button32.BackColor = Color.DarkRed;
            button33.BackColor = Color.DarkSalmon;
            button34.BackColor = Color.DarkSeaGreen;
            button35.BackColor = Color.DarkSlateBlue;
            button36.BackColor = Color.DarkSlateGray;
            button37.BackColor = Color.DarkTurquoise;
            button38.BackColor = Color.DarkViolet;
            button39.BackColor = Color.DeepPink;
            button40.BackColor = Color.DeepSkyBlue;
            button41.BackColor = Color.DimGray;
            button42.BackColor = Color.DodgerBlue;
            button43.BackColor = Color.Firebrick;
            button44.BackColor = Color.FloralWhite;
            button45.BackColor = Color.ForestGreen;
            button46.BackColor = Color.Fuchsia;
            button47.BackColor = Color.Gainsboro;
            button48.BackColor = Color.GhostWhite;
            button49.BackColor = Color.Gold;
            button50.BackColor = Color.Goldenrod;
            button51.BackColor = Color.Gray;
            button52.BackColor = Color.Green;
            button53.BackColor = Color.GreenYellow;
            button54.BackColor = Color.Honeydew;
            button55.BackColor = Color.HotPink;
            button56.BackColor = Color.IndianRed;
            button57.BackColor = Color.Indigo;
            button58.BackColor = Color.Ivory;
            button59.BackColor = Color.Khaki;
            button60.BackColor = Color.Lavender;
            button61.BackColor = Color.LavenderBlush;
            button62.BackColor = Color.LawnGreen;
            button63.BackColor = Color.LemonChiffon;
            button64.BackColor = Color.LightBlue;
            button65.BackColor = Color.LightCoral;
            button66.BackColor = Color.LightCyan;
            button67.BackColor = Color.LightGoldenrodYellow;
            button68.BackColor = Color.LightGray;
            button69.BackColor = Color.LightGreen;
            button70.BackColor = Color.LightPink;
            button71.BackColor = Color.LightSalmon;
            button72.BackColor = Color.LightSeaGreen;
            button73.BackColor = Color.LightSkyBlue;
            button74.BackColor = Color.LightSlateGray;
            button75.BackColor = Color.LightSteelBlue;
            button76.BackColor = Color.LightYellow;
            button77.BackColor = Color.Lime;
            button78.BackColor = Color.LimeGreen;
            button79.BackColor = Color.Linen;
            button80.BackColor = Color.Magenta;
            button81.BackColor = Color.Maroon;
            button82.BackColor = Color.MediumAquamarine;
            button83.BackColor = Color.MediumBlue;
            button84.BackColor = Color.MediumOrchid;
            button85.BackColor = Color.MediumPurple;
            button86.BackColor = Color.MediumSeaGreen;
            button87.BackColor = Color.MediumSlateBlue;
            button88.BackColor = Color.MediumSpringGreen;
            button89.BackColor = Color.MediumTurquoise;
            button90.BackColor = Color.MediumVioletRed;
            button91.BackColor = Color.MidnightBlue;
            button92.BackColor = Color.MintCream;
            button93.BackColor = Color.MistyRose;
            button94.BackColor = Color.Moccasin;
            button95.BackColor = Color.NavajoWhite;
            button96.BackColor = Color.Navy;
            button97.BackColor = Color.OldLace;
            button98.BackColor = Color.Olive;
            button99.BackColor = Color.OliveDrab;
            button100.BackColor = Color.Orange;
            button101.BackColor = Color.OrangeRed;
            button102.BackColor = Color.Orchid;
            button103.BackColor = Color.PaleGoldenrod;
            button104.BackColor = Color.PaleGreen;
            button105.BackColor = Color.PaleTurquoise;
            button106.BackColor = Color.PaleVioletRed;
            button107.BackColor = Color.PapayaWhip;
            button108.BackColor = Color.PeachPuff;
            button109.BackColor = Color.Peru;
            button110.BackColor = Color.Pink;
            button111.BackColor = Color.Plum;
            button112.BackColor = Color.PowderBlue;
            button113.BackColor = Color.Purple;
            button114.BackColor = Color.Red;
            button115.BackColor = Color.RosyBrown;
            button116.BackColor = Color.RoyalBlue;
            button117.BackColor = Color.SaddleBrown;
            button118.BackColor = Color.Salmon;
            button119.BackColor = Color.SandyBrown;
            button120.BackColor = Color.SeaGreen;
            button121.BackColor = Color.SeaShell;
            button122.BackColor = Color.Sienna;
            button123.BackColor = Color.Silver;
            button124.BackColor = Color.SkyBlue;
            button125.BackColor = Color.SlateBlue;
            button126.BackColor = Color.SlateGray;
            button127.BackColor = Color.Snow;
            button128.BackColor = Color.SpringGreen;
            button129.BackColor = Color.SteelBlue;
            button130.BackColor = Color.Tan;
            button131.BackColor = Color.Teal;
            button132.BackColor = Color.Thistle;
            button133.BackColor = Color.Tomato;
            button134.BackColor = Color.Transparent;
            button135.BackColor = Color.Turquoise;
            button136.BackColor = Color.Violet;
            button137.BackColor = Color.Wheat;
            button138.BackColor = Color.White;
            button139.BackColor = Color.WhiteSmoke;
            button140.BackColor = Color.Yellow;
            button141.BackColor = Color.Yellow;
            #endregion

            Button[] all_buttons = new Button[] {
                kalc_0, kalc_1, kalc_2, kalc_3, kalc_4, kalc_5, kalc_6,
                kalc_7, kalc_8, kalc_9, kalc_back, kalc_del, login_button
            };

            for (int i = 0; i < all_buttons.Length; i++)
            {
                all_buttons[i].BackColor = Color.FromArgb(255, 215, 228, 242);
                all_buttons[i].FlatStyle = FlatStyle.Flat;
            }

            this.BackColor = Color.FromArgb(255, 120, 215, 124);

        }

        private void name_color(object sender, EventArgs e)
        {
            MessageBox.Show((sender as Button).BackColor.Name.ToString());
        }

        private void EAN_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                login_button.PerformClick();
        }
    }
}
