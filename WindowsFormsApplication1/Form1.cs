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
using Excel = Microsoft.Office.Interop.Excel;
using System.Threading;

namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        int tek_pol = -1; //текущая категория
        int tek_kat = -1; //текущая подкатегория
        string tek_bl = ""; //текущее блюдо
        int tek_mod = 0; //текущий модификатор
        int cur_row = 0; //выбранная строка
        blydo_1 temp_blydo; //для хранения блюда с модификатором

        MsgBoxYesNo MsgBox; //для окон предупреждения
        public string razreshenie; //чтобы знать кто выдал разрешение на действие

        int vyvod = 0; //сколько элементов выведено на данный момент

        int chek_number = 1; //номер чека
        string oficiant; //имя официанта
        int summa = 0; //сумма заказа

        public List<kategoria_1> new_kategorii = new List<kategoria_1>(); //список катгеорий, их подкатегорий и их блюд
        public List<blydo_1> dont_show_bl = new List<blydo_1>();

        private Button[] buttons; //массив кликабельных кнопок для катгеорий и пр.
        Button[] kalculator; //массив кнопок из калькулятора

        Form mather; //форма авторизации

        string dolznost; //должность сотрудника
        int zakaz; //номер стола

        string time = "СЕЙЧАС"; //время забития заказа

        //это для работы с БД. Всё.
        OleDbCommand komand;
        OleDbDataReader post;
        OleDbConnection podkl = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + @"\\PC-SPB\Users\RESTORAN\baza.mdb");
        //

        bool obz_mod = false;

        public Form1(Form temp1, string temp2, string temp3, string temp4)
        {
            InitializeComponent();
            mather = temp1;
            oficiant = temp2;
            dolznost = temp3;
            zakaz = int.Parse(temp4);
            int h = Screen.PrimaryScreen.Bounds.Height;
            int w = Screen.PrimaryScreen.Bounds.Width;

            this.MaximumSize = new Size(w, h);
            this.MinimumSize = new Size(w, h);
            this.Size = new Size(w, h);
        }
        private void refresh_summa()
        {
            label1.Text = "Сумма заказа: " + summa.ToString();
        }//перезапись суммы
        private void refresh_kat()
        {
            if (vyvod == 0)
                back.Visible = false;
            else
                back.Visible = true;
            next.Visible = true;

            if (tek_pol == -1)
                gotomenu.Visible = false;
            else
                gotomenu.Visible = true;

            back_kat.Visible = false;
            if (tek_pol == -1)
            {
                for (int i = 0; i < buttons.Length; i++)
                {
                    if (vyvod < new_kategorii.Count)
                    {
                        buttons[i].Visible = true;
                        buttons[i].Text = new_kategorii[i].name;
                    }
                    else
                        buttons[i].Visible = false;
                    vyvod++;
                }

                if (vyvod >= new_kategorii.Count)
                    next.Visible = false;

                vyvod -= buttons.Length;
            }
            else
            {
                if (vyvod == 0)
                    back.Visible = false;
                else
                    back.Visible = true;

                for (int i = 0; i < buttons.Length; i++)
                {
                    if (vyvod < new_kategorii[tek_pol].podcategorii.Count)
                    {
                        buttons[i].Visible = true;
                        buttons[i].Text = new_kategorii[tek_pol].podcategorii[vyvod].name;
                    }
                    else
                        buttons[i].Visible = false;
                    vyvod++;
                }

                if (vyvod >= new_kategorii[tek_pol].podcategorii.Count)
                    next.Visible = false;

                vyvod -= buttons.Length;
            }
        } //пролистывние категорий
        private void refresh_b()
        {
            next.Visible = true;
            back_kat.Visible = true;

            var temp = new_kategorii[tek_pol].podcategorii[tek_kat];

            if (vyvod == 0)
                back.Visible = false;
            else
                back.Visible = true;

            for (int i = 0; i < buttons.Length; i++)
            {
                if (vyvod < temp.blyda.Count)
                {
                    buttons[i].Visible = true;
                    buttons[i].Text = temp.blyda[vyvod].name;
                }
                else
                    buttons[i].Visible = false;
                vyvod++;
            }

            if (vyvod >= temp.blyda.Count)
                next.Visible = false;

            vyvod -= buttons.Length;
        } //пролистывание блюд
        private void save_answer()
        {
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                if (dataGridView1.Rows[i].DefaultCellStyle.BackColor == Color.Empty)
                {
                    MsgBox = new MsgBoxYesNo("Сохранить несохранённые изменения?", 2);
                    if (MsgBox.ShowDialog() == DialogResult.Yes)
                    {
                        save.Visible = true;
                        save.PerformClick();
                    }
                    return;
                }
            }
        }//есть ли несохранённые изменения
        private void select_modifikator(int pol, string mod, bool obz) //выбор модификатора
        {
            #region ОБЯЗАТЕЛЬНЫЕ МОДИФИКАТОРЫ
            if (obz_mod)
            {
                string temp_b;
                #region МОДИФИКАТОРЫ - МАССИВ
                if (temp_blydo.modifikatory_ob != null)
                {
                    temp_b = button1.Text;
                    if (pol < temp_blydo.modifikatory_ob.Length)
                    {
                        next.Visible = true;
                        if (vyvod == 0)
                            back.Visible = false;
                        else
                            back.Visible = true;

                        for (int i = 0; i < buttons.Length; i++)
                        {
                            if (vyvod < temp_blydo.modifikatory_ob[pol].Length)
                            {
                                buttons[i].Visible = true;
                                buttons[i].Text = temp_blydo.modifikatory_ob[pol][vyvod];
                            }
                            else
                                buttons[i].Visible = false;
                            vyvod++;
                        }

                        if (vyvod >= temp_blydo.modifikatory_ob[pol].Count())
                            next.Visible = false;

                        vyvod -= buttons.Length;
                    }
                    if ((vyvod > 0) && (temp_b != button1.Text))
                        return;

                    if (pol > 0)
                    {
                        if (tek_bl.IndexOf('@') != -1)
                        {
                            int temp_index = tek_bl.IndexOf('@');
                            string temp_tek_bl = tek_bl;
                            tek_bl = tek_bl.Substring(0, temp_index);
                            tek_bl += mod;
                            tek_bl += temp_tek_bl.Substring(temp_index + 1, temp_tek_bl.Length - temp_index - 1);
                        }
                        else
                        {
                            tek_bl += mod;
                            if (pol < temp_blydo.modifikatory_ob.Length)
                                tek_bl += ", ";
                        }
                    }

                    if (pol == temp_blydo.modifikatory_ob.Length)
                        add_bl(temp_blydo.kod, temp_blydo.cena, tek_bl, true, true);
                }
                #endregion
                #region МОДИФИКАТОРЫ - СПИСОК
                else
                {
                    temp_b = button1.Text;
                    if (pol < temp_blydo.modifikatory_ob_List.Length)
                    {
                        next.Visible = true;
                        if (vyvod == 0)
                            back.Visible = false;
                        else
                            back.Visible = true;

                        for (int i = 0; i < buttons.Length; i++)
                        {
                            if (vyvod < temp_blydo.modifikatory_ob_List[pol].Count)
                            {
                                buttons[i].Visible = true;
                                buttons[i].Text = temp_blydo.modifikatory_ob_List[pol][vyvod].name;
                            }
                            else
                                buttons[i].Visible = false;
                            vyvod++;
                        }

                        if (vyvod >= temp_blydo.modifikatory_ob_List[pol].Count())
                            next.Visible = false;

                        vyvod -= buttons.Length;
                    }

                    if ((vyvod > 0) && (temp_b != button1.Text))
                        return;

                    if (pol > 0)
                    {
                        if (tek_bl.IndexOf('@') != -1)
                        {
                            int temp_index = tek_bl.IndexOf('@');
                            string temp_tek_bl = tek_bl;
                            tek_bl = tek_bl.Substring(0, temp_index);
                            tek_bl += mod;
                            tek_bl += temp_tek_bl.Substring(temp_index + 1, temp_tek_bl.Length - temp_index - 1);
                        }
                        else
                        {
                            tek_bl += mod;
                            if (pol < temp_blydo.modifikatory_ob_List.Length)
                                tek_bl += ", ";
                        }

                        int ti = temp_blydo.modifikatory_ob_List[pol - 1].FindIndex(x => x.name == mod);
                        blydo_1 tb = temp_blydo.modifikatory_ob_List[pol - 1][ti];
                        if (pol != temp_blydo.modifikatory_ob_List.Length)
                            add_bl(tb.kod, tb.cena, tb.name_in_chek, false, false);
                        else
                        {
                            add_bl(tb.kod, tb.cena, tb.name_in_chek, true, true);
                            return;
                        }
                    }
                    else
                        add_bl(temp_blydo.kod, temp_blydo.cena, tek_bl, false, false);
                }
                #endregion
            }
            #endregion
            else
            {
                string temp_b;
                #region МОДИФИКАТОРЫ - МАССИВ
                if (temp_blydo.modifikatory_nob != null)
                {
                    temp_b = button1.Text;
                    if (pol < temp_blydo.modifikatory_nob.Length)
                    {
                        next.Visible = true;
                        if (vyvod == 0)
                            back.Visible = false;
                        else
                            back.Visible = true;

                        for (int i = 0; i < buttons.Length; i++)
                        {
                            if (vyvod < temp_blydo.modifikatory_nob[pol].Length)
                            {
                                buttons[i].Visible = true;
                                buttons[i].Text = temp_blydo.modifikatory_nob[pol][vyvod];
                            }
                            else
                                buttons[i].Visible = false;
                            vyvod++;
                        }

                        if (vyvod >= temp_blydo.modifikatory_nob[pol].Count())
                            next.Visible = false;

                        vyvod -= buttons.Length;
                    }
                    if ((vyvod > 0) && (temp_b != button1.Text))
                        return;

                    dataGridView1.CurrentRow.Cells[1].Value = dataGridView1.CurrentRow.Cells[1].Value.ToString() + ": " + mod;

                    gotomenu.PerformClick();
                }
                #endregion
                #region МОДИФИКАТОРЫ - СПИСОК
                else
                {
                    temp_b = button1.Text;
                    if (pol < temp_blydo.modifikatory_ob_List.Length)
                    {
                        next.Visible = true;
                        if (vyvod == 0)
                            back.Visible = false;
                        else
                            back.Visible = true;

                        for (int i = 0; i < buttons.Length; i++)
                        {
                            if (vyvod < temp_blydo.modifikatory_ob_List[pol].Count)
                            {
                                buttons[i].Visible = true;
                                buttons[i].Text = temp_blydo.modifikatory_ob_List[pol][vyvod].name;
                            }
                            else
                                buttons[i].Visible = false;
                            vyvod++;
                        }

                        if (vyvod >= temp_blydo.modifikatory_ob_List[pol].Count())
                            next.Visible = false;

                        vyvod -= buttons.Length;
                    }

                    if ((vyvod > 0) && (temp_b != button1.Text))
                        return;

                    if (pol > 0)
                    {
                        if (tek_bl.IndexOf('@') != -1)
                        {
                            int temp_index = tek_bl.IndexOf('@');
                            string temp_tek_bl = tek_bl;
                            tek_bl = tek_bl.Substring(0, temp_index);
                            tek_bl += mod;
                            tek_bl += temp_tek_bl.Substring(temp_index + 1, temp_tek_bl.Length - temp_index - 1);
                        }
                        else
                        {
                            tek_bl += mod;
                            if (pol < temp_blydo.modifikatory_ob_List.Length)
                                tek_bl += ", ";
                        }

                        int ti = temp_blydo.modifikatory_ob_List[pol - 1].FindIndex(x => x.name == mod);
                        blydo_1 tb = temp_blydo.modifikatory_ob_List[pol - 1][ti];
                        if (pol != temp_blydo.modifikatory_ob_List.Length)
                            add_bl(tb.kod, tb.cena, tb.name_in_chek, false, false);
                        else
                        {
                            add_bl(tb.kod, tb.cena, tb.name_in_chek, true, true);
                            return;
                        }
                    }
                    else
                        add_bl(temp_blydo.kod, temp_blydo.cena, tek_bl, false, false);
                }
                #endregion
            }
        }
        public void add_bl(int temp_kod, int temp_cena, string temp_name, bool vozvrat, bool chistka) //добавление блюда
        {
            podkl.Open();
            komand = new OleDbCommand("SELECT Count(*) FROM [Состав заказа] WHERE [Код заказа]=" + zakaz.ToString(), podkl);
            post = komand.ExecuteReader();
            post.Read();
            int kolvo_bl = int.Parse(post.GetValue(0).ToString());
            int dob = 1;
            podkl.Close();

            dataGridView1.Rows.Add(temp_kod, temp_name, temp_cena, 1, 1, kolvo_bl + dob);
            dataGridView1.Rows[dataGridView1.Rows.Count - 1].Cells[0].Selected = true;

            DataGridViewRow temp2 = dataGridView1.CurrentRow;
            if (kalc_Text.Text != "")
                temp2.Cells[3].Value = int.Parse(kalc_Text.Text);
            summa += int.Parse(temp2.Cells[2].Value.ToString()) * int.Parse(temp2.Cells[3].Value.ToString());

            refresh_summa();

            if (chistka)
            {
                temp_blydo = null;
                tek_mod = 0;
                tek_bl = "";
                kalc_Text.Text = "";
                for (int ii = 0; ii < kalculator.Length; ii++)
                {
                    kalculator[ii].Visible = true;
                }
                kalc_Text.Visible = true;

                gotomenu.Visible = true;
                back_kat.Visible = true;
                del_poz.Visible = true;
                go.Visible = true;
                go_ochered.Visible = true;
                pozze.Visible = true;
                save.Visible = true;
                ochered.Visible = true;
                search.Visible = true;
                zakryt.Visible = true;
                other_stol.Visible = true;
                block.Visible = true;
                print_predchek.Visible = true;
                del_poz.Visible = true;
            }

            if (vozvrat)
            {
                gotomenu.PerformClick();
                #region быстрый клик
                /*if (textBox1.Text != "")
                    buttons[int.Parse(textBox1.Text)].PerformClick();
                if (checkBox1.Checked)
                    next.PerformClick();
                if (textBox2.Text != "")
                    buttons[int.Parse(textBox2.Text)].PerformClick();
                if (checkBox2.Checked)
                    next.PerformClick();
                if (checkBox3.Checked)
                    next.PerformClick();
                if (textBox3.Text != "")
                    buttons[int.Parse(textBox3.Text)].PerformClick();
                if (checkBox4.Checked)
                    next.PerformClick();
                if (checkBox5.Checked)
                    next.PerformClick();*/
                #endregion
            }

            if (dataGridView1.Rows.Count >= 30)
            {
                dataGridView1.Columns[0].Width = 96;
                dataGridView1.Columns[1].Width = 330;
                dataGridView1.Columns[2].Width = 95;
                dataGridView1.Columns[3].Width = 88;
                dataGridView1.Columns[4].Width = 66;
            }
        }
        private void id_kategoria(object sender, EventArgs e) //тык на кнопку из Buttons
        {
            if (tek_pol == -1) //если не выбрана категория
            {
                tek_pol = new_kategorii.FindIndex(x => x.name == (sender as Button).Text);
                vyvod = 0;
                refresh_kat();
            }
            else
            {
                if (tek_kat == -1) //если не выбрана подкатегория
                {
                    tek_kat = new_kategorii[tek_pol].podcategorii.FindIndex(x => x.name == (sender as Button).Text);
                    vyvod = 0;
                    refresh_b();
                }
                else
                {
                    if (tek_bl == "") //если не выбрано блюдо
                    {
                        temp_blydo = new_kategorii[tek_pol].podcategorii[tek_kat].blyda.Find(x => x.name == (sender as Button).Text);

                        tek_bl = temp_blydo.name;
                        if (temp_blydo.name_in_chek != "")
                            tek_bl = temp_blydo.name_in_chek + " ";

                        if ((temp_blydo.modifikatory_ob != null) || (temp_blydo.modifikatory_ob_List != null)) //если есть модификаторы
                        {
                            tek_bl = temp_blydo.name + ": ";
                            for (int ii = 0; ii < kalculator.Length; ii++)
                            {
                                kalculator[ii].Visible = false;
                            }
                            kalc_Text.Visible = false;

                            gotomenu.Visible = false;
                            back_kat.Visible = false;
                            del_poz.Visible = false;
                            go_ochered.Visible = false;
                            go.Visible = false;
                            pozze.Visible = false;
                            save.Visible = false;
                            ochered.Visible = false;
                            search.Visible = false;
                            zakryt.Visible = false;
                            other_stol.Visible = false;
                            block.Visible = false;
                            print_predchek.Visible = false;

                            vyvod = 0;

                            obz_mod = true;
                            select_modifikator(tek_mod, null, true);
                        }
                        else
                            add_bl(temp_blydo.kod, temp_blydo.cena, tek_bl, false, true);
                    }
                    else //если выбран модификатор
                    {
                        select_modifikator(++tek_mod, (sender as Button).Text, true);
                    }
                }
            }
        }
        private void kalkulator_Click(object sender, EventArgs e)
        {
            if ((sender as Button).Text == "Стол")
            {
                if (otmena.Visible == true)
                {
                    podkl.Open();
                    komand = new OleDbCommand("Select [Код заказа] FROM [Заказы] WHERE [Когда закрыт] Is NULL AND [Номер стола]=" + kalc_Text.Text, podkl);
                    post = komand.ExecuteReader();

                    if (post.Read() == false)
                        MsgBox = new MsgBoxYesNo("Данного стола нет.", 1);
                    else
                    {
                        var a = dataGridView1.CurrentRow;
                        MsgBox = new MsgBoxYesNo("Подтвердите перенос блюда " + a.Cells[1].Value + " в количестве " + a.Cells[3].Value + " на " + kalc_Text.Text + " стол", 2);
                        if (MsgBox.ShowDialog() == DialogResult.Yes)
                        {
                            komand = new OleDbCommand("INSERT INTO [Состав заказа]([Код заказа], [Код блюда], [Наименование], [Количество], [Когда добавлено]) VALUES (" +
                                        post.GetValue(0).ToString() + ", " + dataGridView1[0, a.Index].Value.ToString() + ", '" + dataGridView1[1, a.Index].Value.ToString() + "', " + dataGridView1[3, a.Index].Value.ToString() + ", '" + DateTime.Now.ToString() + "' LIMIT 1)", podkl);
                            post = komand.ExecuteReader();

                            summa -= int.Parse(dataGridView1[2, a.Index].Value.ToString()) * int.Parse(dataGridView1[3, a.Index].Value.ToString());

                            komand = new OleDbCommand("DELETE FROM [Состав заказа] WHERE [Код заказа]=" + zakaz.ToString() + " AND [Код блюда]=" + dataGridView1[0, a.Index].Value.ToString() + " AND [Количество]=" + dataGridView1[3, a.Index].Value.ToString() + " LIMIT 1", podkl);
                            post = komand.ExecuteReader();

                            dataGridView1.Rows.RemoveAt(a.Index);

                            otmena.PerformClick();
                        }
                    }
                    podkl.Close();

                }
            }
            //############################################
            Button temp = (sender as Button);
            switch (temp.Text)
            {
                case "СТЕРЕТЬ":
                    kalc_Text.Text = "";
                    break;
                case "<-":
                    if (kalc_Text.Text != "")
                        kalc_Text.Text = kalc_Text.Text.Substring(0, kalc_Text.Text.Length - 1);
                    break;
                case "Количество":
                    if (dataGridView1.Rows.Count != 0)
                    {
                        DataGridViewRow temp2 = dataGridView1.CurrentRow;
                        if ((temp2.DefaultCellStyle.BackColor == Color.SpringGreen) || (temp2.DefaultCellStyle.BackColor == Color.Yellow))
                            return;
                        if (temp2.DefaultCellStyle.BackColor == Color.LightGray)
                        {
                            if ((dolznost == "ОФ") || (dolznost == "CO"))
                            {
                                MsgBox = new MsgBoxYesNo("Недостаточно прав. Необходимо подтверждение менеджера.", 1);
                                MsgBox.Owner = this;
                                if (MsgBox.ShowDialog() != DialogResult.Yes)
                                    return;
                            }
                        }

                        #region ДЛЯ ОБЫЧНЫХ

                        //if ()
                        podkl.Open();
                        komand = new OleDbCommand("UPDATE [Состав заказа] SET [Количество]=" + kalc_Text.Text + " WHERE [Код заказа]=" + zakaz.ToString() + " AND [Номер добавления]=" + temp2.Cells[5].Value.ToString(), podkl);
                        post = komand.ExecuteReader();
                        podkl.Close();

                        summa -= int.Parse(temp2.Cells[2].Value.ToString()) * int.Parse(temp2.Cells[3].Value.ToString());
                        summa += int.Parse(temp2.Cells[2].Value.ToString()) * int.Parse(kalc_Text.Text);
                        dataGridView1[3, dataGridView1.CurrentRow.Index].Value = int.Parse(kalc_Text.Text);
                        kalc_del.PerformClick();
                        refresh_summa();
                        #endregion
                    }
                    break;
                default:
                    kalc_Text.Text += temp.Text;
                    break;
            }

        }
        private void Form1_Load(object sender, EventArgs e)
        {
            #region ИНИЦИАЛИЗАЦИЯ ВСЕГО
            buttons = new Button[] { button1, button2, button3, button4, button5, button6, button7, button8, button9, button10, button11, button12 };
            kalculator = new Button[] { kalc_0, kalc_1, kalc_2, kalc_3, kalc_4, kalc_5, kalc_6, kalc_7, kalc_8, kalc_9, kalc_back, kalc_del, kolichestvo, kod_bl };
            Button[] all_buttons = new Button[] {
                button1, button2, button3, button4, button5, button6, button7,
                button8, button9, button10, button11,button12,
                kalc_0, kalc_1, kalc_2, kalc_3, kalc_4, kalc_5, kalc_6,
                kalc_7, kalc_8, kalc_9, kalc_back, kalc_del, kolichestvo, kod_bl,
                gotomenu, back_kat, del_poz, back, next, search, otmena,
                go, pozze, save, other_stol, block, print_predchek, per_poz,
                ochered, go_ochered
            };
            #endregion
            #region РАСКРАСКА
            for (int i = 0; i < all_buttons.Length; i++)
            {
                all_buttons[i].BackColor = Color.FromArgb(255, 215, 228, 242);
                all_buttons[i].FlatStyle = FlatStyle.Flat;
            }
            for (int i = 0; i < buttons.Length; i++)
            {
                buttons[i].BackColor = Color.Plum;
            }
            for (int i = 0; i < kalculator.Length; i++)
            {
                kalculator[i].BackColor = Color.DarkKhaki;
            }

            kod_bl.BackColor = Color.Peru;
            pozze.BackColor = Color.Orange;
            save.BackColor = Color.LightSkyBlue;
            search.BackColor = Color.SteelBlue;
            next.BackColor = Color.Thistle;
            back.BackColor = Color.Thistle;
            go_ochered.BackColor = Color.Firebrick;

            this.BackColor = Color.FromArgb(255, 120, 215, 124);
            del_poz.BackColor = Color.Red;
            #endregion
            #region ДЛЯ ПЕРЕНОСА ЗАКАЗА
            if (zakaz == 0)
            {
                del_poz.Visible = false;
                per_poz.Visible = true;

                for (int i = 0; i < buttons.Length; i++)
                {
                    buttons[i].Visible = false;
                }
                next.Visible = false;
                back.Visible = false;
                del_poz.Visible = true;

                for (int i = 0; i < kalculator.Length; i++)
                {
                    kalculator[i].Visible = false;
                }
                kalc_Text.Visible = false;

                search.Visible = false;
                print_predchek.Visible = false;
                save.Visible = false;
                pozze.Visible = false;
                go.Visible = false;
            }
            else
            {
                del_poz.Visible = true;
                per_poz.Visible = false;
            }
            #endregion
            #region ОСНОВНЫЕ КАТЕГОРИИ
            podkl.Open();
            komand = new OleDbCommand("Select [Название], [Код] From [Коды] WHERE [Код]<10 ORDER BY [Код]", podkl);
            post = komand.ExecuteReader();

            string[] temp = new string[10];

            while (post.Read())
            {
                string name = post.GetValue(0).ToString();
                int kod = int.Parse(post.GetValue(1).ToString());

                new_kategorii.Add(new kategoria_1(name, kod));
            }
            #endregion
            #region ПОДКАТЕГОРИИ
            for (int j = 1; j <= new_kategorii.Count; j++)
            {
                komand = new OleDbCommand("Select [Название], [Код] From [Коды] WHERE [Код] LIKE '" + j.ToString() + "__' ORDER BY [Код]", podkl);
                post = komand.ExecuteReader();

                temp = new string[100];

                while (post.Read())
                {
                    string name = post.GetValue(0).ToString();
                    int kod = int.Parse(post.GetValue(1).ToString());

                    new_kategorii[j - 1].podcategorii.Add(new podcategroia_1(name, kod));
                }
            }
            #endregion
            #region БЛЮДА
            DateTime temp_date = DateTime.Now;
            string date_now = temp_date.Month.ToString() + "/" + temp_date.Day.ToString() + "/" + temp_date.Year + " " + temp_date.ToLongTimeString();
            for (int i = 0; i < new_kategorii.Count; i++)
            {
                for (int j = 0; j < new_kategorii[i].podcategorii.Count; j++)
                {
                    OleDbCommand komand2;
                    OleDbDataReader post2;
                    komand2 = new OleDbCommand("SELECT [Код блюда], [Наименование], [Наименование в чеке], [Цена], [Обязательные модификаторы], [Необязательные модификаторы] FROM [Блюда] WHERE [Код блюда] LIKE '" + new_kategorii[i].podcategorii[j].kod.ToString() + "____' AND ([Дата вывода]>#" + date_now + "# OR [Дата вывода] IS Null) ORDER BY [Код блюда]", podkl);
                    post2 = komand2.ExecuteReader();

                    while (post2.Read())
                    {
                        int kod = int.Parse(post2.GetValue(0).ToString());
                        string naimenovanie = post2.GetValue(1).ToString();
                        string naimenovanie_chek = post2.GetValue(2).ToString();
                        int cena = int.Parse(post2.GetValue(3).ToString());
                        string modifikator_ob = post2.GetValue(4).ToString();
                        string modifikator_nob = post2.GetValue(5).ToString();

                        new_kategorii[i].podcategorii[j].blyda.Add(new blydo_1(kod, naimenovanie, naimenovanie_chek, cena, modifikator_ob, modifikator_nob, podkl));
                    }
                    post2.Close();

                }
            }

            komand = new OleDbCommand("SELECT [Код блюда], [Наименование], [Наименование в чеке], [Цена], [Обязательные модификаторы], [Необязательные модификаторы] FROM [Блюда] WHERE [Код блюда]<1000000 AND ([Дата вывода]>#" + date_now + "# OR [Дата вывода] IS Null) ORDER BY [Код блюда]", podkl);
            post = komand.ExecuteReader();

            while (post.Read())
            {
                int kod = int.Parse(post.GetValue(0).ToString());
                string naimenovanie = post.GetValue(1).ToString();
                string naimenovanie_chek = post.GetValue(2).ToString();
                int cena = int.Parse(post.GetValue(3).ToString());
                string modifikator_ob = post.GetValue(4).ToString();
                string modifikator_nob = post.GetValue(5).ToString();

                dont_show_bl.Add(new blydo_1(kod, naimenovanie, naimenovanie_chek, cena, modifikator_ob, modifikator_nob, podkl));
            }
            #endregion
            #region ЗАПОЛНЕНИЕ ТАБЛИЦЫ
            dataGridView1.BackgroundColor = Color.White;
            komand = new OleDbCommand("Select [Код блюда], [Количество], [Наименование], [Когда добавлено], [Готовить позже], [Номер в очереди], [Номер добавления] FROM [Состав заказа] WHERE [Код заказа]=" + zakaz.ToString() + " ORDER BY [Номер добавления]", podkl);
            post = komand.ExecuteReader();

            DateTime data_select = default(DateTime);

            while (post.Read())
            {
                int temp_kod = int.Parse(post.GetValue(0).ToString());
                int temp_kolvo = int.Parse(post.GetValue(1).ToString());
                string temp_name = post.GetValue(2).ToString();
                int temp_ochered = int.Parse(post.GetValue(5).ToString());
                int temp_dob = int.Parse(post.GetValue(6).ToString());

                if ((data_select != Convert.ToDateTime(post.GetValue(3).ToString())) || (data_select == default(DateTime)))
                {
                    data_select = Convert.ToDateTime(post.GetValue(3).ToString());
                    dataGridView1.Rows.Add(null, data_select.ToString(), null, null, null);
                    dataGridView1.Rows[dataGridView1.Rows.Count - 1].DefaultCellStyle.BackColor = Color.SpringGreen;
                }

                OleDbCommand komand2 = new OleDbCommand("SELECT [Цена] FROM [Блюда] WHERE [Код блюда]=" + temp_kod.ToString(), podkl);
                OleDbDataReader post2 = komand2.ExecuteReader();
                post2.Read();

                summa += int.Parse(post2.GetValue(0).ToString()) * temp_kolvo;

                dataGridView1.Rows.Add(temp_kod, temp_name, int.Parse(post2.GetValue(0).ToString()), temp_kolvo, temp_ochered, temp_dob);
                if (post.GetValue(4).ToString() == "True")
                {
                    dataGridView1.Rows[dataGridView1.Rows.Count - 1].DefaultCellStyle.BackColor = Color.Yellow;
                }
            }
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                if ((dataGridView1.Rows[i].DefaultCellStyle.BackColor != Color.SpringGreen) && (dataGridView1.Rows[i].DefaultCellStyle.BackColor != Color.Yellow))
                    dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.LightGray;
            }
            dataGridView1.Rows.Add(null, time, null, null);
            dataGridView1.Rows[dataGridView1.Rows.Count - 1].DefaultCellStyle.BackColor = Color.SpringGreen;

            dataGridView1.ClearSelection();

            refresh_summa();
            podkl.Close();


            if (dataGridView1.Rows.Count >= 30)
            {
                dataGridView1.Columns[0].Width = 96;
                dataGridView1.Columns[1].Width = 330;
                dataGridView1.Columns[2].Width = 95;
                dataGridView1.Columns[3].Width = 88;
                dataGridView1.Columns[4].Width = 66;
            }

            #endregion

            if (zakaz == 0)
                return;
            else
                refresh_kat();
        }
        private void next_Click(object sender, EventArgs e)
        {
            vyvod += buttons.Length;
            if (tek_kat == -1)
                refresh_kat();
            else
            {
                if (tek_bl == "")
                    refresh_b();
                else
                {
                    obz_mod = true;
                    select_modifikator(tek_mod, null, true);
                }
            }
        }
        private void back_Click(object sender, EventArgs e)
        {
            vyvod -= buttons.Length;
            if (tek_kat == -1)
                refresh_kat();
            else
            {
                if (tek_bl == "")
                    refresh_b();
                else
                {
                    obz_mod = true;
                    select_modifikator(tek_mod, null, true);
                }
            }
        }
        private void gotomenu_Click(object sender, EventArgs e)
        {
            tek_pol = -1;
            tek_kat = -1;
            vyvod = 0;
            temp_blydo = null;
            tek_mod = 0;
            tek_bl = "";
            refresh_kat();
        }
        private void print_predchek_Click(object sender, EventArgs e)
        {
            save_answer();
            if (summa == 0)
            {
                MsgBox = new MsgBoxYesNo("Нельзя выводить на печать чек с нулевой суммой", 1);
                MsgBox.ShowDialog();
                return;
            }
            int r = 3;

            if ((dolznost == "ОФ") || (dolznost == "СО"))
            {
                MsgBox = new MsgBoxYesNo("Недостаточно прав. Необходимо подтверждение менеджера.", 1);
                MsgBox.Owner = this;
                if (MsgBox.ShowDialog() != DialogResult.Yes)
                    return;
            }
            else
            {
                razreshenie = (Owner as select_stol).kod;
            }

            podkl.Open();
            komand = new OleDbCommand("SELECT Count(*) FROM [Предчеки]", podkl);
            post = komand.ExecuteReader();
            post.Read();
            komand = new OleDbCommand("INSERT INTO [Предчеки]([Код предчека], [Заказ], [Когда напечатан], [Кто разрешил]) VALUES (" +
                            (int.Parse(post.GetValue(0).ToString()) + 1).ToString() + ", " + zakaz.ToString() + ", '" + DateTime.Now.ToString() + "', '" + razreshenie + "')", podkl);

            post = komand.ExecuteReader();
            podkl.Close();

            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook workBook = excelApp.Workbooks.Add(Type.Missing);
            Excel.Worksheet workSheet = (Excel.Worksheet)excelApp.Worksheets.get_Item(1);

            //excelApp.DisplayAlerts = false;
            excelApp.Visible = false;

            Excel.Range temp;

            temp = workSheet.get_Range("A1", "G2");
            temp.Merge(Type.Missing);
            temp.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            temp.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            temp.Value2 = "ООО «Воробушек»";
            temp.Font.Name = "Calibri";
            temp.Font.Size = 14;
            temp.Font.Bold = true;
            temp.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlDouble;

            int count_bl = 0;

            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                if (dataGridView1.Rows[i].DefaultCellStyle.BackColor == Color.LightGray)
                {
                    DataGridViewRow temp2 = dataGridView1.Rows[i];
                    workSheet.Cells[count_bl + 3, 1] = temp2.Cells[1].Value.ToString();
                    workSheet.Cells[count_bl + 3, 6] = "*" + temp2.Cells[3].Value.ToString();
                    workSheet.Cells[count_bl + 3, 7] = "=" + (int.Parse(temp2.Cells[2].Value.ToString()) * int.Parse(temp2.Cells[3].Value.ToString())).ToString();
                    //#############################################################
                    workSheet.Cells[count_bl + 3, 1].Font.Name = "Calibri";
                    workSheet.Cells[count_bl + 3, 1].Font.Size = 8;
                    workSheet.Cells[count_bl + 3, 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                    workSheet.Cells[count_bl + 3, 1].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;

                    workSheet.Cells[count_bl + 3, 6].Font.Name = "Calibri";
                    workSheet.Cells[count_bl + 3, 6].Font.Size = 8;
                    workSheet.Cells[count_bl + 3, 6].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                    workSheet.Cells[count_bl + 3, 6].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;

                    workSheet.Cells[count_bl + 3, 7].Font.Name = "Calibri";
                    workSheet.Cells[count_bl + 3, 7].Font.Size = 8;
                    workSheet.Cells[count_bl + 3, 7].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                    workSheet.Cells[count_bl + 3, 7].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    //#############################################################
                    r++;
                    count_bl++;
                }
            }

            for (int i = 1; i <= 8; i++)
            {
                workSheet.Columns[i].ColumnWidth = 3.86;
            }

            r++;
            //#############################################################
            temp = workSheet.get_Range("A" + r.ToString(), "E" + (r + 1).ToString());
            temp.Merge(Type.Missing);
            temp.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
            temp.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            temp.Value2 = "Итого к оплате";
            temp.Font.Name = "Calibri";
            temp.Font.Size = 8;
            temp.Font.Bold = true;

            temp = workSheet.get_Range("F" + r.ToString(), "G" + (r + 1).ToString());
            temp.Merge(Type.Missing);
            temp.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
            temp.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            temp.Value2 = summa.ToString();
            temp.Font.Name = "Calibri";
            temp.Font.Size = 8;
            temp.Font.Bold = true;
            //#############################################################
            r += 2;

            //#############################################################
            temp = workSheet.get_Range("A" + (r).ToString(), "G" + (r + 4).ToString());
            temp.Font.Name = "Calibri";
            temp.Font.Size = 6;

            //#############################################################
            //#############################################################
            workSheet.Cells[++r, 1] = "ОФИЦИАНТ";
            temp = workSheet.get_Range("C" + r.ToString(), "G" + r.ToString());
            temp.Merge(Type.Missing);
            temp.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
            temp.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            temp.Value2 = oficiant;

            temp = workSheet.get_Range("A" + r.ToString(), "G" + r.ToString());
            temp.Borders.get_Item(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous;
            //#############################################################

            workSheet.Cells[++r, 1] = "ДАТА";
            temp = workSheet.get_Range("B" + r.ToString(), "C" + r.ToString());
            temp.Merge(Type.Missing);
            temp.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
            temp.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            temp.Value2 = DateTime.Now.ToShortDateString();
            workSheet.Cells[r, 4] = "ВРЕМЯ";
            workSheet.Cells[r, 5] = DateTime.Now.ToShortTimeString();
            workSheet.Cells[r, 5].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
            //#############################################################

            workSheet.Cells[++r, 1] = "КАССА";
            temp = workSheet.get_Range("B" + r.ToString(), "G" + r.ToString());
            temp.Merge(Type.Missing);
            temp.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
            temp.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            temp.Value2 = "1";
            //#############################################################

            workSheet.Cells[++r, 1] = "ЧЕК";
            temp = workSheet.get_Range("B" + r.ToString(), "G" + r.ToString());
            temp.Merge(Type.Missing);
            temp.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
            temp.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            temp.Value2 = (chek_number++).ToString();
            //#############################################################
            r += 2;
            temp = workSheet.get_Range("A" + r.ToString(), "G" + (r + 1).ToString());
            temp.Merge(Type.Missing);
            temp.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            temp.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            //temp.Value2 = "Ресторан Воробушек - забота жареного барашка";
            temp.Value2 = "ЯВЛЯЕТСЯ ПРЕЧЕКОМ";
            temp.Font.Bold = true;
            temp.Font.Italic = true;
            temp.Font.Size = 14;
            temp.WrapText = true;
            //#############################################################
            r += 2;
            temp = workSheet.get_Range("A" + r.ToString(), "G" + (r + 1).ToString());
            temp.Merge(Type.Missing);
            temp.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            temp.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            //temp.Value2 = "Ресторан Воробушек - забота жареного барашка";
            temp.Value2 = "Вознаграждение официанту приветсвуется, но всегда остаётся на Ваше усмотрение.";
            temp.Font.Italic = true;
            temp.Font.Size = 7;
            temp.WrapText = true;

            temp = workSheet.get_Range("A1", "G" + (r + 1).ToString());
            temp.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            temp.Borders.get_Item(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlDot;
            temp.Borders.get_Item(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlDot;


            workSheet.PageSetup.LeftMargin = 0;
            workSheet.PageSetup.RightMargin = 0;
            workSheet.PageSetup.TopMargin = 0;
            workSheet.PageSetup.BottomMargin = 0;
            workSheet.PageSetup.HeaderMargin = 0;
            workSheet.PageSetup.FooterMargin = 0;
            /*
            workSheet.PrintOut();
            GC.Collect();
            GC.WaitForPendingFinalizers();

            workBook.Close(false, Type.Missing, Type.Missing);

            excelApp.Quit();*/
            excelApp.Visible = true;
        }
        private void back_kat_Click(object sender, EventArgs e)
        {
            temp_blydo = null;
            tek_mod = 0;
            tek_bl = "";
            tek_kat = -1;
            vyvod = 0;
            refresh_kat();
        }
        private void del_poz_Click(object sender, EventArgs e)
        {
            if (tek_bl == "")
            {
                if (dataGridView1.Rows.Count != 0)
                {
                    int a = dataGridView1.CurrentRow.Index;

                    if (dataGridView1.Rows[a].DefaultCellStyle.BackColor == Color.SpringGreen)
                        return;

                    razreshenie = (Owner as select_stol).kod;
                    if (dataGridView1.Rows[a].DefaultCellStyle.BackColor != Color.Empty)
                    {
                        if ((dolznost == "ОФ") || (dolznost == "СО"))
                        {
                            MsgBox = new MsgBoxYesNo("Недостаточно прав. Необходимо подтверждение менеджера.", 1);
                            MsgBox.Owner = this;
                            if (MsgBox.ShowDialog() != DialogResult.Yes)
                                return;
                        }
                    }

                    podkl.Open();

                    if (dataGridView1.Rows[a].DefaultCellStyle.BackColor == Color.Empty)
                    {
                        komand = new OleDbCommand("INSERT INTO [Удаленные блюда]([Код заказа], [Код блюда], [Наименование], [Количество], [Кем удалено], [Когда удалено], [Причина удаления]) VALUES (" +
                            zakaz.ToString() + ", " + dataGridView1[0, a].Value.ToString() + ", '" + dataGridView1[1, a].Value.ToString() + "', " + dataGridView1[3, a].Value.ToString() + ", '" + razreshenie + "', '" + DateTime.Now.ToString() + "', 'ОО')", podkl);
                        post = komand.ExecuteReader();

                        summa -= int.Parse(dataGridView1[2, a].Value.ToString()) * int.Parse(dataGridView1[3, a].Value.ToString());

                        komand = new OleDbCommand("DELETE FROM [Состав заказа] WHERE [Код заказа]=" + zakaz.ToString() + " AND [Код блюда]=" + dataGridView1[0, a].Value.ToString() + " AND [Количество]=" + dataGridView1[3, a].Value.ToString(), podkl);
                        post = komand.ExecuteReader();
                        podkl.Close();

                        dataGridView1.Rows.RemoveAt(a);
                        refresh_summa();
                        return;
                    }


                    string prichina = "ОО";
                    switch (new MsgBoxYesNo("Причина удаления", 3).ShowDialog())
                    {
                        case (DialogResult.Cancel):
                            prichina = "ОГ";
                            break;
                        case (DialogResult.Retry):
                            prichina = "СЛ";
                            break;
                        case (DialogResult.Ignore):
                            prichina = "НП";
                            break;
                    }

                    switch (new MsgBoxYesNo("Выберите действие", 4).ShowDialog())
                    {
                        case (DialogResult.No):
                            return;
                        case (DialogResult.Yes):
                            komand = new OleDbCommand("INSERT INTO [Удаленные блюда]([Код заказа], [Код блюда], [Наименование], [Количество], [Кем удалено], [Когда удалено], [Причина удаления]) VALUES (" +
                                zakaz.ToString() + ", " + dataGridView1[0, a].Value.ToString() + ", '" + dataGridView1[1, a].Value.ToString() + "', " + dataGridView1[3, a].Value.ToString() + ", '" + razreshenie + "', '" + DateTime.Now.ToString() + "', '" + prichina + "')", podkl);
                            post = komand.ExecuteReader();
                            break;
                        case (DialogResult.OK):
                            komand = new OleDbCommand("INSERT INTO [Состав заказа]([Код заказа], [Код блюда], [Наименование], [Количество], [Когда добавлено]) VALUES (" +
                                "0, " + dataGridView1[0, a].Value.ToString() + ", '" + dataGridView1[1, a].Value.ToString() + "', " + dataGridView1[3, a].Value.ToString() + ", '" + DateTime.Now.ToString() + "')", podkl);
                            post = komand.ExecuteReader();
                            break;

                    }
                    summa -= int.Parse(dataGridView1[2, a].Value.ToString()) * int.Parse(dataGridView1[3, a].Value.ToString());


                    komand = new OleDbCommand("DELETE FROM [Состав заказа] WHERE [Код заказа]=" + zakaz.ToString() + " AND [Код блюда]=" + dataGridView1[0, a].Value.ToString() + " AND [Количество]=" + dataGridView1[3, a].Value.ToString(), podkl);
                    post = komand.ExecuteReader();
                    podkl.Close();

                    dataGridView1.Rows.RemoveAt(a);

                    refresh_summa();
                }
            }
        }
        private void save_Click(object sender, EventArgs e)
        {
            podkl.Open();
            time = DateTime.Now.ToString();

            komand = new OleDbCommand("SELECT Count(*) FROM [Состав заказа] WHERE [Код заказа]=" + zakaz.ToString(), podkl);
            post = komand.ExecuteReader();
            post.Read();
            int kolvo_bl = int.Parse(post.GetValue(0).ToString());
            int dob = 1;
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                if ((dataGridView1.Rows[i].DefaultCellStyle.BackColor != Color.LightGray) && (dataGridView1.Rows[i].DefaultCellStyle.BackColor != Color.SpringGreen) && (dataGridView1.Rows[i].DefaultCellStyle.BackColor != Color.Yellow))
                {
                    komand = new OleDbCommand("INSERT INTO [Состав заказа]([Код заказа], [Код блюда], [Наименование], [Количество], [Когда добавлено], [Номер в очереди], [Номер добавления], [Номер в очереди (архив)]) VALUES (" + zakaz.ToString() + ", " + dataGridView1[0, i].Value.ToString() + ", '" + dataGridView1[1, i].Value.ToString() + "', " + dataGridView1[3, i].Value.ToString() + ", '" + time + "', " + dataGridView1[4, i].Value.ToString() + ", " + (kolvo_bl + dob).ToString() + ", " + dataGridView1[4, i].Value.ToString() + ")", podkl);
                    post = komand.ExecuteReader();
                    dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.LightGray;
                    dob++;
                }
                if (dataGridView1[1, i].Value.ToString() == "СЕЙЧАС")
                {
                    dataGridView1[1, i].Value = time;
                }
            }
            podkl.Close();

            time = DateTime.Now.ToString();
            dataGridView1.Rows.Add(null, "СЕЙЧАС", null, null);
            dataGridView1.Rows[dataGridView1.Rows.Count - 1].DefaultCellStyle.BackColor = Color.SpringGreen;

            dataGridView1.ClearSelection();
        }
        private void block_Click(object sender, EventArgs e)
        {

            save_answer();
            mather.Show();
            this.Close();
        }
        private void other_stol_Click(object sender, EventArgs e)
        {
            save_answer();
            Owner.Show();
            this.Close();
        }
        private void go_Click(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count != 0)
            {
                int a = dataGridView1.CurrentRow.Index;
                if (dataGridView1.Rows[a].DefaultCellStyle.BackColor == Color.Empty)
                {
                    podkl.Open();
                    komand = new OleDbCommand("INSERT INTO [Состав заказа]([Код заказа], [Код блюда], [Наименование], [Количество], [Когда добавлено]) VALUES (" + zakaz.ToString() + ", " + dataGridView1[0, a].Value.ToString() + ", '" + dataGridView1[1, a].Value.ToString() + "', " + dataGridView1[3, a].Value.ToString() + ", '" + DateTime.Now.ToString() + "')", podkl);
                    post = komand.ExecuteReader();
                    podkl.Close();
                    dataGridView1.Rows[a].DefaultCellStyle.BackColor = Color.LightGray;
                }
                if (dataGridView1.Rows[a].DefaultCellStyle.BackColor == Color.Yellow)
                {
                    var temp2 = dataGridView1.CurrentRow;
                    podkl.Open();
                    komand = new OleDbCommand("UPDATE [Состав заказа] SET [Готовить позже]=FALSE WHERE [Код заказа]=" + zakaz.ToString() + " AND [Код блюда]=" + temp2.Cells[0].Value.ToString() + " AND [Количество]=" + temp2.Cells[3].Value.ToString() + " AND [Готовить позже]=TRUE", podkl);
                    post = komand.ExecuteReader();
                    podkl.Close();
                    dataGridView1.Rows[a].DefaultCellStyle.BackColor = Color.LightGray;
                }
            }
        }
        private void pozze_Click(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count != 0)
            {
                int a = dataGridView1.CurrentRow.Index;
                if (dataGridView1.Rows[a].DefaultCellStyle.BackColor == Color.Empty)
                {
                    podkl.Open();
                    komand = new OleDbCommand("SELECT Count(*) FROM [Состав заказа] WHERE [Код заказа]=" + zakaz.ToString(), podkl);
                    post = komand.ExecuteReader();
                    post.Read();
                    int kolvo_bl = int.Parse(post.GetValue(0).ToString());
                    int dob = 1;
                    komand = new OleDbCommand("INSERT INTO [Состав заказа]([Код заказа], [Код блюда], [Наименование], [Количество], [Когда добавлено], [Готовить позже], [Номер в очереди], [Номер добавления], [Номер в очереди (архив)]) VALUES (" + zakaz.ToString() + ", " + dataGridView1[0, a].Value.ToString() + ", '" + dataGridView1[1, a].Value.ToString() + "', " + dataGridView1[3, a].Value.ToString() + ", '" + DateTime.Now.ToString() + "', TRUE, " + dataGridView1[4, a].Value.ToString() + ", " + (kolvo_bl + dob).ToString() + ", " + dataGridView1[4, a].Value.ToString() + ")", podkl);
                    post = komand.ExecuteReader();
                    podkl.Close();
                    dataGridView1.Rows[a].DefaultCellStyle.BackColor = Color.Yellow;
                }
            }
        }
        private void kod_bl_Click(object sender, EventArgs e)
        {
            podkl.Open();
            komand = new OleDbCommand("Select [Код блюда], [Наименование], [Наименование в чеке], [Цена], [Обязательные модификаторы], [Необязательные модификаторы], [Необязательные модификаторы] From [Блюда] WHERE [Код блюда]=" + kalc_Text.Text, podkl);
            post = komand.ExecuteReader();
            post.Read();

            int kod_2 = int.Parse(post.GetValue(0).ToString());
            string naimenovanie_2 = post.GetValue(1).ToString();
            string naimenovanie_chek_2 = post.GetValue(2).ToString();
            int cena_2 = int.Parse(post.GetValue(3).ToString());
            string modifikator_ob_2 = post.GetValue(4).ToString();
            string modifikator_nob_2 = post.GetValue(5).ToString();


            temp_blydo = new blydo_1(kod_2, naimenovanie_2, naimenovanie_chek_2, cena_2, modifikator_ob_2, modifikator_nob_2, podkl);
            kalc_Text.Text = "";
            podkl.Close();

            tek_bl = temp_blydo.name;

            if (temp_blydo.name_in_chek != "")
                tek_bl = temp_blydo.name_in_chek + " ";

            if ((temp_blydo.modifikatory_ob != null) || (temp_blydo.modifikatory_ob_List != null)) //если есть модификаторы
            {
                if (temp_blydo.name_in_chek == "")
                    tek_bl = temp_blydo.name + ": ";

                for (int ii = 0; ii < kalculator.Length; ii++)
                {
                    kalculator[ii].Visible = false;
                }
                kalc_Text.Visible = false;

                gotomenu.Visible = false;
                back_kat.Visible = false;
                del_poz.Visible = false;
                go_ochered.Visible = false;
                go.Visible = false;
                pozze.Visible = false;
                save.Visible = false;
                ochered.Visible = false;
                search.Visible = false;
                zakryt.Visible = false;
                other_stol.Visible = false;
                block.Visible = false;
                print_predchek.Visible = false;

                vyvod = 0;

                tek_kat = -2;
                tek_pol = -2;
                obz_mod = true;
                select_modifikator(tek_mod, null, true);
            }
            else
                add_bl(temp_blydo.kod, temp_blydo.cena, tek_bl, true, true);
        }
        private void search_Click(object sender, EventArgs e)
        {
            search frm = new search();
            frm.Owner = this;
            frm.Show();
            this.Hide();
        }
        private void dataGridView1_Click(object sender, EventArgs e)
        {
            if (cur_row == dataGridView1.CurrentRow.Index)
            {
                dataGridView1.ClearSelection();
                cur_row = -1;
            }
            else
                cur_row = dataGridView1.CurrentRow.Index;
        }
        private void per_poz_Click(object sender, EventArgs e)
        {
            var a = dataGridView1.CurrentRow;
            if ((a.Index != 0) && (a.DefaultCellStyle.BackColor == Color.LightGray))
            {
                per_poz.Visible = false;
                otmena.Visible = true;
                for (int i = 0; i < kalculator.Length; i++)
                {
                    kalculator[i].Visible = true;
                }
                kalc_Text.Visible = true;
                kod_bl.Visible = false;
                kolichestvo.Text = "Стол";
            }
        }
        private void otmena_Click(object sender, EventArgs e)
        {
            per_poz.Visible = true;
            otmena.Visible = false;
            for (int i = 0; i < kalculator.Length; i++)
            {
                kalculator[i].Visible = false;
            }
            kalc_Text.Visible = false;
        }
        private void ochered_Click(object sender, EventArgs e)
        {
            if ((dataGridView1.Rows.Count != 0) && (kalc_Text.Text != ""))
            {
                DataGridViewRow temp2 = dataGridView1.CurrentRow;
                if (temp2.DefaultCellStyle.BackColor == Color.LightGray)
                {
                    if ((dolznost == "ОФ") || (dolznost == "CO"))
                    {
                        MsgBox = new MsgBoxYesNo("Недостаточно прав. Необходимо подтверждение менеджера.", 1);
                        MsgBox.Owner = this;
                        if (MsgBox.ShowDialog() != DialogResult.Yes)
                            return;
                    }

                    podkl.Open();
                    komand = new OleDbCommand("UPDATE [Состав заказа] SET [Номер в очереди]=" + kalc_Text.Text + " WHERE [Код заказа]=" + zakaz.ToString() + " AND [Номер добавления]=" + temp2.Cells[5].Value.ToString(), podkl);
                    post = komand.ExecuteReader();
                    podkl.Close();
                }
                summa -= int.Parse(temp2.Cells[2].Value.ToString()) * int.Parse(temp2.Cells[3].Value.ToString());
                summa += int.Parse(temp2.Cells[2].Value.ToString()) * int.Parse(kalc_Text.Text);
                dataGridView1[4, dataGridView1.CurrentRow.Index].Value = int.Parse(kalc_Text.Text);
                kalc_del.PerformClick();
                refresh_summa();
            }
        }
        private void button13_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                var temp2 = dataGridView1.Rows[i];
                if (temp2.DefaultCellStyle.BackColor == Color.Yellow)
                {
                    MsgBox = new MsgBoxYesNo("Поданы не все блюда. Отметьте их как готовить сейчас.", 1);
                    MsgBox.ShowDialog();
                    return;
                }
            }

            MsgBox = new MsgBoxYesNo("Подтвердите закрытие стола.", 2);
            if (MsgBox.ShowDialog() == DialogResult.Yes)
            {
                refresh_summa();

                DateTime temp_date = DateTime.Now;
                string date_now = temp_date.Month.ToString() + "/" + temp_date.Day.ToString() + "/" + temp_date.Year + " " + temp_date.ToLongTimeString();

                podkl.Open();
                komand = new OleDbCommand("UPDATE [Заказы] SET [Когда закрыт]='" + date_now + "' WHERE [Код заказа]=" + zakaz.ToString(), podkl);
                post = komand.ExecuteReader();
                komand = new OleDbCommand("UPDATE [Заказы] SET [Сумма]=" + summa.ToString() + " WHERE [Код заказа]=" + zakaz.ToString(), podkl);
                post = komand.ExecuteReader();
                podkl.Close();
                other_stol.PerformClick();
            }
        }
        private void go_ochered_Click(object sender, EventArgs e)
        {
            if (kalc_Text.Text != "")
            {
                podkl.Open();
                komand = new OleDbCommand("SELECT MAX([Номер в очереди]) FROM [Состав заказа]", podkl);
                post = komand.ExecuteReader();
                post.Read();

                int max_ochered = int.Parse(post.GetValue(0).ToString());
                MsgBox = new MsgBoxYesNo("Подтвердите запуск очереди", 2);
                if (MsgBox.ShowDialog() == DialogResult.Yes)
                {
                    komand = new OleDbCommand("UPDATE [Состав заказа] SET [Номер в очереди]=1 WHERE [Номер в очереди]=" + kalc_Text.Text, podkl);
                    post = komand.ExecuteReader();
                    for (int i = int.Parse(kalc_Text.Text) + 1; i <= max_ochered; i++)
                    {
                        komand = new OleDbCommand("UPDATE [Состав заказа] SET [Номер в очереди]=" + (i - 1).ToString() + " WHERE [Номер в очереди]=" + i.ToString(), podkl);
                        post = komand.ExecuteReader();
                    }

                    for (int i = 0; i < dataGridView1.Rows.Count; i++)
                    {
                        if (dataGridView1.Rows[i].DefaultCellStyle.BackColor == Color.LightGray)
                        {
                            if (int.Parse(dataGridView1.Rows[i].Cells[4].Value.ToString()) == int.Parse(kalc_Text.Text))
                                dataGridView1.Rows[i].Cells[4].Value = 1;
                            if (int.Parse(dataGridView1.Rows[i].Cells[4].Value.ToString()) > int.Parse(kalc_Text.Text))
                            {
                                dataGridView1.Rows[i].Cells[4].Value = int.Parse(dataGridView1.Rows[i].Cells[4].Value.ToString()) - 1;
                            }
                        }
                    }
                }
                podkl.Close();
            }
        }
        private void kalc_Text_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                if ((kalc_Text.Text.ToLower() == "d1") || (kalc_Text.Text.ToLower() == "в1"))
                {
                    kalc_Text.Text = "";
                    return;
                }

                if (kalc_Text.Text.Length > 8)
                {
                    kalc_Text.Text = "";
                    return;
                }

                kod_bl.PerformClick();
            }
        }
        private void Form1_Activated(object sender, EventArgs e)
        {
            #region ИНИЦИАЛИЗАЦИЯ ВСЕГО
            buttons = new Button[] { button1, button2, button3, button4, button5, button6, button7, button8, button9, button10, button11, button12 };
            kalculator = new Button[] { kalc_0, kalc_1, kalc_2, kalc_3, kalc_4, kalc_5, kalc_6, kalc_7, kalc_8, kalc_9, kalc_back, kalc_del, kolichestvo, kod_bl };
            Button[] all_buttons = new Button[] {
                button1, button2, button3, button4, button5, button6, button7,
                button8, button9, button10, button11,button12,
                kalc_0, kalc_1, kalc_2, kalc_3, kalc_4, kalc_5, kalc_6,
                kalc_7, kalc_8, kalc_9, kalc_back, kalc_del, kolichestvo, kod_bl,
                gotomenu, back_kat, del_poz, back, next, search, otmena,
                go, pozze, save, other_stol, block, print_predchek, per_poz,
                ochered, go_ochered
            };
            #endregion
            #region ПРОРИСОВКА
            this.Invalidate();
            for (int i = 0; i < all_buttons.Length; i++)
            {
                all_buttons[i].Invalidate();
            }
            kalc_Text.Invalidate();
            dataGridView1.Invalidate();
            #endregion
        }

        private void modifikatory_Click(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count != 0)
            {
                DataGridViewRow temp2 = dataGridView1.CurrentRow;

                if ((temp2.DefaultCellStyle.BackColor == Color.SpringGreen) || (temp2.DefaultCellStyle.BackColor == Color.Yellow))
                    return;

                if (temp2.DefaultCellStyle.BackColor == Color.LightGray)
                {
                    if ((dolznost == "ОФ") || (dolznost == "CO"))
                    {
                        MsgBox = new MsgBoxYesNo("Недостаточно прав. Необходимо подтверждение менеджера.", 1);
                        MsgBox.Owner = this;
                        if (MsgBox.ShowDialog() != DialogResult.Yes)
                            return;
                    }
                }
                
                temp_blydo = null;
                int i = 0, j = 0;
                
                while ((temp_blydo == null) && (i < new_kategorii.Count))
                {
                    while (j < new_kategorii[i].podcategorii.Count)
                    {
                        if (new_kategorii[i].podcategorii[j].blyda.FindIndex(x => x.kod.ToString() == temp2.Cells[0].Value.ToString()) == -1)
                            j++;
                        else
                            temp_blydo = new_kategorii[i].podcategorii[j].blyda.Find(x => x.kod.ToString() == temp2.Cells[0].Value.ToString());
                    }

                    j = 0;
                    i++;
                }

                if (temp_blydo == null)
                    temp_blydo = dont_show_bl.Find(x => x.kod.ToString() == temp2.Cells[0].Value.ToString());

                tek_pol = -2;
                tek_kat = -2;
                tek_bl = " ";
                var a = "";
                obz_mod = false;
                select_modifikator(tek_mod, null, false);
            }

        }
    }
}
