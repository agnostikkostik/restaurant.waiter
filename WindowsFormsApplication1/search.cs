using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApplication1
{
    public partial class search : Form
    {
        Form1 mather;
        List<kategoria_1> blyda;
        List<blydo_1> search_result;

        public search()
        {
            InitializeComponent();
        }

        Button[] buttons;
        int vyvod = 0;

        private void search_Load(object sender, EventArgs e)
        {
            mather = this.Owner as Form1;
            blyda = mather.new_kategorii;

            Button[] all_buttons = new Button[] {
                s_0, s_1, s_2, s_3, s_4, s_5, s_6, s_7, s_8, s_9,
                //0   1    2    3    4    5    6    7    8    9
                s_a, s_b, s_v, s_g, s_d, s_ye, s_zh, s_z, s_i,
                //а   б    в    г    д    е     ж     з    и
                s_y, s_k, s_l, s_m, s_n, s_o, s_p, s_r, s_s,
                //й   к    л    м    н    о    п    р    с
                s_t,s_u, s_f, s_h, s_c, s_ch, s_sh, s_shch, s_tz,
                //т  у    ф    х    ц    ч     ш     щ       ъ
                s_yi, s_mz, s_e, s_yu, s_ya, s_backspace, s_back,
                //ы    ь     э    ю     я
                button1, button2, button3, button4, button5, button6, button7,
                button8, button9, button10, button11,button12, button13,
                next, back, exit
            };

            for (int i = 0; i < all_buttons.Length; i++)
            {
                all_buttons[i].BackColor = Color.FromArgb(255, 215, 228, 242);
                all_buttons[i].FlatStyle = FlatStyle.Flat;
            }

            this.BackColor = Color.FromArgb(255, 120, 215, 124);

            buttons = new Button[] {
                button1, button2, button3, button4, button5, button6, button7,
                button8, button9, button10, button11,button12, button13
            };

            search_result = new List<blydo_1>();
        }

        private void s_Click(object sender, EventArgs e)
        {
            Button temp = sender as Button;
            s_Text.Text += temp.Text;
        }

        private void s_back_Click(object sender, EventArgs e)
        {
            if (s_Text.Text != "")
                s_Text.Text = s_Text.Text.Substring(0, s_Text.Text.Length - 1);
        }

        private void s_Text_TextChanged(object sender, EventArgs e)
        {
            vyvod = 0;
            search_result.Clear();

            for (int i = 0; i < blyda.Count; i++)
            {
                for (int j = 0; j < blyda[i].podcategorii.Count; j++)
                {
                    for (int k = 0; k < blyda[i].podcategorii[j].blyda.Count; k++)
                    {
                        if (blyda[i].podcategorii[j].blyda[k].name.ToLower().IndexOf(s_Text.Text.ToLower()) != -1)
                            search_result.Add(blyda[i].podcategorii[j].blyda[k]);
                        if (blyda[i].podcategorii[j].blyda[k].kod.ToString().IndexOf(s_Text.Text) != -1)
                            search_result.Add(blyda[i].podcategorii[j].blyda[k]);
                    }
                }
            }

            refresh();
        }

        private void exit_Click(object sender, EventArgs e)
        {
            mather.Show();
            this.Close();
        }

        private void refresh()
        {
            if (vyvod == 0)
                back.Visible = false;
            else
                back.Visible = true;
            next.Visible = true;

            for (int i = 0; i < buttons.Length; i++)
            {
                if (vyvod < search_result.Count)
                {
                    buttons[i].Visible = true;
                    buttons[i].Text = search_result[vyvod].name;
                }
                else
                    buttons[i].Visible = false;
                vyvod++;
            }

            if (vyvod >= search_result.Count)
                next.Visible = false;

            vyvod -= buttons.Length;
        }

        private void back_Click(object sender, EventArgs e)
        {
            vyvod -= buttons.Length;
            refresh();
        }

        private void next_Click(object sender, EventArgs e)
        {
            vyvod += buttons.Length;
            refresh();
        }

        private void buttons_Click(object sender, EventArgs e)
        {
            blydo_1 temp = search_result.Find(t => t.name == (sender as Button).Text);

            mather.Show();

            mather.kalc_Text.Visible = true;
            mather.kod_bl.Visible = true;
            mather.kalc_Text.Text = temp.kod.ToString();
            mather.kod_bl.PerformClick();

            this.Close();
        }
    }
}
