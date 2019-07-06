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
using System.Threading;
using Excel = Microsoft.Office.Interop.Excel;

namespace WindowsFormsApplication1
{
    public partial class select_stol : Form
    {
        OleDbCommand komand;
        OleDbDataReader post;
        OleDbConnection podkl = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + @"\\PC-SPB\Users\RESTORAN\baza.mdb");


        int vyvod = 0;
        public string kod;
        public string name;
        string dolznost;
        Button[] buttons;
        Button[] kalculator;
        Button[] other;
        List<int> stoly = new List<int>();
        List<int> all_stolyy = new List<int>();

        public select_stol(string temp1, string temp2, string temp3)
        {
            InitializeComponent();
            kod = temp1;
            dolznost = temp2;
            name = temp3;


            int h = Screen.PrimaryScreen.Bounds.Height;
            int w = Screen.PrimaryScreen.Bounds.Width;

            this.MaximumSize = new Size(w, h);
            this.MinimumSize = new Size(w, h);
            this.Size = new Size(w, h);

            #region РАЗМЕРЫ КНОПОК
            Button[] all_buttons = new Button[] {
                kalc_0, kalc_1, kalc_2, kalc_3, kalc_4, kalc_5, kalc_6,
                kalc_7, kalc_8, kalc_9, kalc_back, kalc_del, open,
                button1, button2, button3, button4, button5, button6,
                back, next, block, all_stoly, all_blyda, shutdown,
                refresh_stoly, otcht, perenos, edit_stop
            };
            double temp_w, temp_h;
            for (int i = 0; i < all_buttons.Length; i++)
            {
                temp_w = all_buttons[i].Size.Width / 1280.0;
                temp_h = all_buttons[i].Size.Height / 720.0;
                all_buttons[i].Size = new Size(
                    Convert.ToInt32(w * temp_w),
                    Convert.ToInt32(h * temp_h)
                    );
            }
            #endregion

            #region РАСПОЛОЖЕНИЕ КНОПОК СТОЛОВ
            all_buttons = new Button[] {
                back, button2, button4, button6
            };

            for (int i = 1; i < all_buttons.Length; i++)
            {
                all_buttons[i].Location = new Point(
                    all_buttons[i].Location.X,
                    all_buttons[i- 1].Size.Height + all_buttons[i - 1].Location.Y + 6);
            }

            button1.Location = new Point(back.Location.X + back.Size.Width + 6, button1.Location.Y);
            
            all_buttons = new Button[] {
                button1, button3, button5, next
            };

            for (int i = 1; i < all_buttons.Length; i++)
            {
                all_buttons[i].Location = new Point(
                    button1.Location.X,
                    all_buttons[i].Size.Height + all_buttons[i - 1].Location.Y + 6);
            }
            #endregion

            #region ИЗМЕНЕНИЕ ТЕКСТОБОКСА
            kalc_Text.Location = new Point(
                Convert.ToInt32(button1.Location.X * (741.0 / 481.0)),
                kalc_Text.Location.Y);
            kalc_Text.Font = new Font(kalc_Text.Font.FontFamily,
                float.Parse((((249 * h) / 720.0) * (8.25 / 249)).ToString()),
                kalc_Text.Font.Style);

            temp_w = kalc_Text.Size.Width / 1280.0;
            temp_h = kalc_Text.Size.Height / 720.0;
            kalc_Text.Size = new Size(
                Convert.ToInt32(w * temp_w),
                kalc_Text.Size.Height);
            #endregion

            #region РАСПОЛОЖЕНИЕ КНОПОК КАЛЬКУЛЯТОРА
            #region 1 СТОЛБЕЦ
            kalc_1.Location = new Point(
                kalc_Text.Location.X,
                kalc_Text.Location.Y + kalc_Text.Size.Height + 4);

            all_buttons = new Button[] {
                kalc_1, kalc_4, kalc_7, kalc_del, open
            };

            for (int i = 1; i < all_buttons.Length; i++)
            {
                all_buttons[i].Location = new Point(
                    kalc_1.Location.X,
                    all_buttons[i - 1].Size.Height + all_buttons[i - 1].Location.Y + 6);
            }
            #endregion
            #region 2 СТОЛБЕЦ
            kalc_2.Location = new Point(
                kalc_1.Location.X + kalc_1.Size.Width + 6,
                kalc_1.Location.Y);

            all_buttons = new Button[] {
                kalc_2, kalc_5, kalc_8, kalc_0
            };

            for (int i = 1; i < all_buttons.Length; i++)
            {
                all_buttons[i].Location = new Point(
                    kalc_2.Location.X,
                    all_buttons[i - 1].Size.Height + all_buttons[i - 1].Location.Y + 6);
            }
            #endregion
            #region 3 СТОЛБЕЦ
            kalc_3.Location = new Point(
                kalc_2.Location.X + kalc_2.Size.Width + 6,
                kalc_1.Location.Y);

            all_buttons = new Button[] {
                kalc_3, kalc_6, kalc_9, kalc_back
            };

            for (int i = 1; i < all_buttons.Length; i++)
            {
                all_buttons[i].Location = new Point(
                    kalc_3.Location.X,
                    all_buttons[i - 1].Size.Height + all_buttons[i - 1].Location.Y + 6);
            }
            #endregion
            #endregion

            #region ОТЧЁТЫ И ПОД НИМИ
            otcht.Location = new Point(
                111,
                Convert.ToInt32((button6.Location.Y * 565.0) / 423.0));
            
            all_stoly.Location = new Point(
                otcht.Location.X,
                otcht.Location.Y + otcht.Size.Height + 6);

            all_buttons = new Button[] {
                all_stoly, all_blyda, shutdown,
                refresh_stoly, perenos, edit_stop, block};

            for (int i = 1; i < all_buttons.Length; i++)
            {
                all_buttons[i].Location = new Point(
                    all_buttons[i - 1].Size.Width + all_buttons[i - 1].Location.X + 6,
                    all_stoly.Location.Y);
            }

            #endregion

            #region ПОЛОСА ЗАГРУЗКИ

            temp_w = progressBar1.Size.Width / 1280.0;
            temp_h = progressBar1.Size.Height / 720.0;
            progressBar1.Size = new Size(
                Convert.ToInt32(w * temp_w),
                Convert.ToInt32(h * temp_h)
                );

            progressBar1.Location = new Point(
                12,
                all_stoly.Location.Y + all_stoly.Size.Height + 6);
            #endregion
        }

        public void refresh()
        {
            back.Visible = true;
            next.Visible = true;

            if (stoly == null)
            {
                for (int i = 0; i < buttons.Length; i++)
                    buttons[i].Visible = false;
                back.Visible = false;
                next.Visible = false;
            }
            else
            {
                if (vyvod == 0)
                    back.Visible = false;

                for (int i = 0; i < buttons.Length; i++)
                    buttons[i].Visible = true;

                for (int i = 0; i < buttons.Length; i++)
                {
                    if (vyvod < stoly.Count)
                        buttons[i].Text = stoly[vyvod].ToString();
                    else
                        buttons[i].Visible = false;
                    vyvod++;
                }

                if (vyvod > (stoly.Count - 1))
                    next.Visible = false;

                vyvod -= buttons.Length;
            }
        }

        private void select_stol_Load(object sender, EventArgs e)
        {
            this.Invalidate();
            Button[] all_buttons = new Button[] {
                kalc_0, kalc_1, kalc_2, kalc_3, kalc_4, kalc_5, kalc_6,
                kalc_7, kalc_8, kalc_9, kalc_back, kalc_del, open,
                button1, button2, button3, button4, button5, button6,
                back, next, block, all_stoly, all_blyda, shutdown,
                refresh_stoly, otcht, perenos, edit_stop
            };

            for (int i = 0; i < all_buttons.Length; i++)
            {
                all_buttons[i].Invalidate();
            }
            stoly.Clear();
            kalc_Text.Text = "";


            kalculator = new Button[] {
                kalc_0, kalc_1, kalc_2, kalc_3, kalc_4, kalc_5, kalc_6,
                kalc_7, kalc_8, kalc_9, kalc_back, kalc_del
            };

            other = new Button[] {
                button1, button2, button3, button4, button5, button6,
                back, next,
            };

            for (int i = 0; i < all_buttons.Length; i++)
            {
                all_buttons[i].BackColor = Color.FromArgb(255, 215, 228, 242);
                all_buttons[i].FlatStyle = FlatStyle.Flat;
            }
            this.BackColor = Color.FromArgb(255, 120, 215, 124);


            podkl.Open();
            buttons = new Button[] { button1, button2, button3, button4, button5, button6 };

            komand = new OleDbCommand("Select [Номер стола] FROM [Заказы] WHERE [Код сотрудника]='" + kod + "' AND [Номер стола]<999999 AND [Когда закрыт] Is NULL", podkl);
            post = komand.ExecuteReader();
            while (post.Read())
            {
                stoly.Add(int.Parse(post.GetValue(0).ToString()));
            }
            podkl.Close();

            if ((dolznost == "ОФ") || (dolznost == "СО"))
            {
                all_blyda.Visible = false;
                all_stoly.Visible = false;
                shutdown.Visible = false;
                otcht.Visible = false;
                perenos.Visible = false;
            }
            if (dolznost == "ОФ")
                edit_stop.Visible = false;
            refresh();
        }

        private void open_Click(object sender, EventArgs e)
        {
            if (kalc_Text.Visible == false)
            {
                for (int i = 0; i < kalculator.Length; i++)
                    kalculator[i].Visible = true;
                kalc_Text.Visible = true;
                //for (int i = 0; i < other.Length; i++)
                //    other[i].Visible = false;
            }
            else
            {
                if (kalc_Text.Text == "")
                    return;
                all_stolyy.Clear();
                podkl.Open();
                komand = new OleDbCommand("Select [Номер стола] FROM [Заказы] WHERE [Когда закрыт] Is NULL", podkl);
                post = komand.ExecuteReader();

                while (post.Read())
                {
                    all_stolyy.Add(int.Parse(post.GetValue(0).ToString()));
                }

                if (all_stolyy.FindIndex(x => x == int.Parse(kalc_Text.Text)) != -1)
                    MessageBox.Show("Стол уже открыт!");
                else
                {
                    komand = new OleDbCommand("SELECT Count(*) From [Заказы]", podkl);
                    post = komand.ExecuteReader();
                    post.Read();
                    komand = new OleDbCommand("INSERT INTO [Заказы]([Код заказа], [Код сотрудника], [Номер стола], [Когда открыт]) VALUES (" + post.GetValue(0).ToString() + ", '" + kod + "', " + kalc_Text.Text + ", '" + DateTime.Now.ToString() + "')", podkl);
                    post = komand.ExecuteReader();
                    stoly.Add(int.Parse(kalc_Text.Text));
                }
                podkl.Close();
                button1.Text = kalc_Text.Text;
                button1.Visible = true;
                button1.PerformClick();

                for (int i = 0; i < kalculator.Length; i++)
                    kalculator[i].Visible = false;
                kalc_Text.Visible = false;
                for (int i = 0; i < other.Length; i++)
                    other[i].Visible = true;
                refresh();
            }
        }

        private void all_blyda_Click(object sender, EventArgs e)
        {
            var a = MessageBox.Show("Подтвердите действие.", "", MessageBoxButtons.YesNo);
            if (a == DialogResult.Yes)
            {
                all_bl();
            }
        }

        private void all_bl()
        {
            progressBar1.Visible = true;
            OleDbCommand komand2;
            OleDbDataReader post2;
            OleDbConnection podkl2 = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + @"\\192.168.1.33\Users\RESTORAN\baza.mdb");
            int kolvo;
            double progress = 0;

            podkl2.Open();

            komand2 = new OleDbCommand("Select COUNT(*) FROM [Блюда]", podkl2);
            post2 = komand2.ExecuteReader();
            post2.Read();
            kolvo = int.Parse(post2.GetValue(0).ToString());

            komand2 = new OleDbCommand("Select [Код блюда], [Наименование], [Цена], [Модификаторы] FROM [Блюда] ORDER BY [Код блюда]", podkl2);
            post2 = komand2.ExecuteReader();

            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook workBook = excelApp.Workbooks.Add(Type.Missing);
            Excel.Worksheet workSheet = (Excel.Worksheet)excelApp.Worksheets.get_Item(1); ;

            excelApp.Visible = false;

            int r = 1;
            int g = 1;
            string modifikator = "";
            string[] modifikatory;

            Excel.Range temp;

            while (post2.Read())
            {
                progress += (2 / 3.0) / kolvo * 100;

                temp = workSheet.get_Range("A" + (r).ToString(), "A" + (r + 1).ToString());
                temp.Merge(Type.Missing);
                temp.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                temp.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                temp.Font.Name = "Calibri";
                temp.Font.Size = 8;
                temp.Font.Bold = true;
                temp.WrapText = true;
                temp.Value2 = post2.GetValue(0).ToString().Substring(0, 4) + " " + post2.GetValue(0).ToString().Substring(4, 3);

                temp = workSheet.get_Range("B" + (r).ToString(), "F" + (r + 1).ToString());
                temp.Merge(Type.Missing);
                temp.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                temp.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                temp.Font.Name = "Calibri";
                temp.Font.Size = 7.5;
                temp.Font.Bold = true;
                temp.WrapText = true;
                temp.Value2 = post2.GetValue(1).ToString();

                temp = workSheet.get_Range("G" + (r).ToString(), "G" + (r + 1).ToString());
                temp.Merge(Type.Missing);
                temp.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                temp.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                temp.Font.Name = "Calibri";
                temp.Font.Size = 8;
                temp.Font.Bold = true;
                temp.WrapText = true;
                temp.Value2 = post2.GetValue(2).ToString();

                r += 2;

                modifikator = post2.GetValue(3).ToString();
                if (modifikator != "")
                {
                    string[] temp1 = modifikator.Split(';');
                    modifikatory = new string[temp1.Length];

                    for (int i = 0; i < temp1.Length; i++)
                    {
                        temp = workSheet.get_Range("A" + (r).ToString(), "G" + (r + 1).ToString());
                        temp.Merge(Type.Missing);
                        temp.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        temp.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                        temp.Font.Name = "Calibri";
                        temp.Font.Size = 15;
                        temp.Font.Bold = true;
                        temp.WrapText = true;
                        temp.Value2 = "МОДИФИКАТОР " + (i + 1).ToString();
                        r += 2;

                        string[] temp2 = temp1[i].Split(',');
                        for (int j = 0; j < temp2.Length; j++)
                        {
                            temp = workSheet.get_Range("A" + (r).ToString(), "G" + (r + 1).ToString());
                            temp.Merge(Type.Missing);
                            temp.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                            temp.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                            temp.Font.Name = "Calibri";
                            temp.Font.Size = 12;
                            temp.Font.Bold = false;
                            temp.WrapText = true;
                            temp.Value2 = temp2[j];
                            r += 2;
                        }

                    }
                }

                temp = workSheet.get_Range("A" + (g).ToString(), "G" + (r - 1).ToString());
                temp.Select();
                temp.Borders.get_Item(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous;
                temp.Borders.get_Item(Excel.XlBordersIndex.xlEdgeLeft).Weight = Excel.XlBorderWeight.xlMedium;
                temp.Borders.get_Item(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous;
                temp.Borders.get_Item(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlMedium;
                temp.Borders.get_Item(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous;
                temp.Borders.get_Item(Excel.XlBordersIndex.xlEdgeTop).Weight = Excel.XlBorderWeight.xlMedium;
                temp.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous;
                temp.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).Weight = Excel.XlBorderWeight.xlMedium;
                g = r;
                progressBar1.Value = Convert.ToInt32(progress);
            }
            for (int i = 1; i <= 7; i++)
            {
                workSheet.Columns[i].ColumnWidth = 3.86;
            }
            for (int i = 1; i <= r; i++)
            {
                progress += (1 / 6.0) / kolvo * 100;
                workSheet.Rows[i].RowHeight = 11.5;
                workSheet.Cells[i, 1].Select();
            }
            if (progress >= 100)
                progressBar1.Value = 100;
            else
                progressBar1.Value = Convert.ToInt32(progress);

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
            progressBar1.Visible = false;
            progressBar1.Value = 0;
        }

        private void kalc_Click(object sender, EventArgs e)
        {
            Button temp = (sender as Button);
            switch (temp.Text)
            {
                case "DEL":
                    kalc_Text.Text = "";
                    break;
                case "<-":
                    if (kalc_Text.Text != "")
                        kalc_Text.Text = kalc_Text.Text.Substring(0, kalc_Text.Text.Length - 1);
                    break;
                default:
                    kalc_Text.Text += temp.Text;
                    break;
            }
        }

        private void next_Click(object sender, EventArgs e)
        {
            vyvod += buttons.Length;
            refresh();
        }

        private void back_Click(object sender, EventArgs e)
        {
            vyvod -= buttons.Length;
            refresh();
        }

        private void refresh_stoly_Click(object sender, EventArgs e)
        {
            podkl.Close();
            all_stolyy.Clear();
            podkl = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + @"\\PC-SPB\Users\RESTORAN\baza.mdb");
        }

        private void stol_Click(object sender, EventArgs e)
        {
            Button temp = sender as Button;
            podkl.Open();
            komand = new OleDbCommand("Select [Код заказа] FROM [Заказы] WHERE [Когда закрыт] Is NULL AND [Номер стола]=" + temp.Text, podkl);
            post = komand.ExecuteReader();
            post.Read();
            Form1 frm = new WindowsFormsApplication1.Form1(Owner, name, dolznost, post.GetValue(0).ToString());
            frm.Owner = this;
            frm.Show();
            podkl.Close();
            this.Hide();
        }

        private void all_stoly_Click(object sender, EventArgs e)
        {
            all_stolyy.Clear();
            podkl.Open();
            komand = new OleDbCommand("Select [Номер стола] FROM [Заказы] WHERE [Когда закрыт] Is NULL AND [Номер стола]<999999", podkl);
            post = komand.ExecuteReader();

            while (post.Read())
            {
                all_stolyy.Add(int.Parse(post.GetValue(0).ToString()));
            }

            podkl.Close();
            stoly = all_stolyy;
            vyvod = 0;
            refresh();
        }

        private void block_Click(object sender, EventArgs e)
        {
            Owner.Show();
            this.Close();
        }

        private void shutdown_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void otcht_Click(object sender, EventArgs e)
        {
            othcety frm = new othcety();
            frm.Owner = this;
            frm.Show();
            this.Hide();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            Button temp = sender as Button;
            podkl.Open();
            komand = new OleDbCommand("Select [Код заказа] FROM [Заказы] WHERE [Когда закрыт] Is NULL AND [Номер стола]=999999", podkl);
            post = komand.ExecuteReader();
            post.Read();
            Form1 frm = new WindowsFormsApplication1.Form1(Owner, name, dolznost, post.GetValue(0).ToString());
            frm.Owner = this;
            frm.Show();
            podkl.Close();
            this.Hide();
        }

        private void select_stol_Activated(object sender, EventArgs e)
        {
            
        }
    }
}
