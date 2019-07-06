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

namespace WindowsFormsApplication1
{
    public partial class othcety : Form
    {
        public othcety()
        {
            InitializeComponent();
        }

        Button[] buttons;
        List<string> oficianty = new List<string>();
        List<int> vyr_oficianty = new List<int>();

        int vyvod = 0;
        int tek_pol;

        OleDbCommand komand;
        OleDbDataReader post;
        OleDbConnection podkl = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + @"\\PC-SPB\Users\RESTORAN\baza.mdb");

        private void hide_butons()
        {
            for (int i = 0; i < buttons.Length; i++)
            {
                buttons[i].Visible = false;
            }
            next.Visible = false;
            back.Visible = false;
        }

        private void show_buttons()
        {
            for (int i = 0; i < buttons.Length; i++)
            {
                buttons[i].Visible = true;
            }
            next.Visible = true;
            back.Visible = true;
        }

        private void refresh()
        {
            if (vyvod == 0)
                back.Visible = false;
            else
                back.Visible = true;
            next.Visible = true;

            switch (tek_pol)
            {
                case (1):
                    for (int i = 0; i < buttons.Length; i++)
                    {
                        if (vyvod < oficianty.Count)
                        {
                            buttons[i].Visible = true;
                            buttons[i].Text = oficianty[vyvod];
                        }
                        else
                            buttons[i].Visible = false;
                        vyvod++;
                    }

                    if (vyvod >= oficianty.Count)
                        next.Visible = false;

                    vyvod -= buttons.Length;
                    break;
            }
        }

        private void othcety_Load(object sender, EventArgs e)
        {
            Button[] all_buttons = new Button[] {
                vyr_all, vyr_of,
                del_bl, bron, oper, skidki, skidki_all,
                ending_bl, ban_bl, rashod, exit,
                button1, button2, button3, button4, button5, button6, button7,
                button8, button9, button10, button11,button12, button13, button14,
                next, back
            };

            buttons = new Button[] {
                button1, button2, button3, button4, button5, button6, button7,
                button8, button9, button10, button11,button12, button13, button14
            };

            for (int i = 0; i < all_buttons.Length; i++)
            {
                all_buttons[i].BackColor = Color.FromArgb(255, 215, 228, 242);
                all_buttons[i].FlatStyle = FlatStyle.Flat;
            }
            this.BackColor = Color.FromArgb(255, 120, 215, 124);

            podkl.Open();
            komand = new OleDbCommand("SELECT [Код сотрудника], Sum([Сумма]) AS [Sum-Сумма] FROM [Заказы] GROUP BY [Код сотрудника]", podkl);
            post = komand.ExecuteReader();
            while (post.Read())
            {
                oficianty.Add(post.GetValue(0).ToString());
                if (post.GetValue(1).ToString() != "")
                    vyr_oficianty.Add(int.Parse(post.GetValue(1).ToString()));
                else
                    vyr_oficianty.Add(0);
            }
            podkl.Close();
        }

        private void exit_Click(object sender, EventArgs e)
        {
            Owner.Show();
            this.Close();
        }

        private void del_bl_Click(object sender, EventArgs e)
        {
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook workBook = excelApp.Workbooks.Add(Type.Missing);
            Excel.Worksheet workSheet = (Excel.Worksheet)excelApp.Worksheets.get_Item(1);

            excelApp.Visible = false;

            Excel.Range temp;
            int r = 7;

            temp = workSheet.get_Range("A1", "H2");
            temp.Merge(Type.Missing);
            temp.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            temp.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            temp.Value2 = "ООО «Воробушек»";
            temp.Font.Name = "Calibri";
            temp.Font.Size = 14;
            temp.Font.Bold = true;
            temp.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlDouble;

            temp = workSheet.get_Range("A3", "H3");
            temp.Merge(Type.Missing);
            temp.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            temp.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            temp.Value2 = "Расход блюд";
            temp.Font.Name = "Calibri";
            temp.Font.Size = 9;

            temp = workSheet.get_Range("A4", "H4");
            temp.Merge(Type.Missing);
            temp.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            temp.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            temp.Value2 = "Кассовый день " + DateTime.Now.ToShortDateString().ToString();
            temp.Font.Name = "Calibri";
            temp.Font.Size = 9;

            temp = workSheet.get_Range("A5", "H5");
            temp.Merge(Type.Missing);
            temp.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            temp.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            temp.Value2 = DateTime.Now.ToString();
            temp.Font.Name = "Calibri";
            temp.Font.Size = 9;
            temp.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlDash;

            temp = workSheet.get_Range("A6", "D6");
            temp.Merge(Type.Missing);
            temp.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            temp.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            temp.Value2 = "Блюдо";
            temp.Font.Name = "Calibri";
            temp.Font.Size = 9;
            temp.Font.Bold = true;
            temp.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlDash;

            temp = workSheet.get_Range("E6", "F6");
            temp.Merge(Type.Missing);
            temp.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            temp.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            temp.Value2 = "Кол-во";
            temp.Font.Name = "Calibri";
            temp.Font.Size = 9;
            temp.Font.Bold = true;
            temp.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlDash;

            temp = workSheet.get_Range("G6", "H6");
            temp.Merge(Type.Missing);
            temp.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            temp.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            temp.Value2 = "Сумма";
            temp.Font.Name = "Calibri";
            temp.Font.Size = 9;
            temp.Font.Bold = true;
            temp.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlDash;

            podkl.Open();
            komand = new OleDbCommand("SELECT [Блюда].[Наименование], Sum([Удаленные блюда].[Количество]) AS [Sum - Количество], Sum([Блюда].[Цена]) AS[Sum - Цена], [Удаленные блюда].[Кем удалено], [Удаленные блюда].[Причина удаления], [Заказы].[Код сотрудника] FROM [Заказы] INNER JOIN ([Блюда] INNER JOIN [Удаленные блюда] ON [Блюда].[Код блюда] = [Удаленные блюда].[Код блюда]) ON [Заказы].[Код заказа] = [Удаленные блюда].[Код заказа] GROUP BY  [Удаленные блюда].[Когда удалено], [Удаленные блюда].[Код блюда], [Блюда].[Наименование], [Удаленные блюда].[Кем удалено], [Удаленные блюда].[Причина удаления], [Заказы].[Код сотрудника]", podkl);
            post = komand.ExecuteReader();
            while (post.Read())
            {
                temp = workSheet.get_Range("A" + r.ToString(), "D" + r.ToString());
                temp.Merge(Type.Missing);
                temp.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                temp.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                temp.Value2 = post.GetValue(0).ToString();
                temp.Font.Name = "Calibri";
                temp.Font.Size = 7;

                temp = workSheet.get_Range("E" + r.ToString(), "F" + r.ToString());
                temp.Merge(Type.Missing);
                temp.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                temp.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                temp.Value2 = post.GetValue(1).ToString();
                temp.Font.Name = "Calibri";
                temp.Font.Size = 7;

                temp = workSheet.get_Range("G" + r.ToString(), "H" + r.ToString());
                temp.Merge(Type.Missing);
                temp.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                temp.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                temp.Value2 = post.GetValue(2).ToString();
                temp.Font.Name = "Calibri";
                temp.Font.Size = 7;
                r++;

                OleDbCommand give_me_name;
                OleDbDataReader post_temp;

                give_me_name = new OleDbCommand("SELECT [Фамилия], [Имя] FROM [Сотрудники] WHERE [id сотрудника]='" + post.GetValue(5).ToString() + "'", podkl);
                post_temp = give_me_name.ExecuteReader();
                post_temp.Read();

                temp = workSheet.get_Range("A" + r.ToString(), "B" + r.ToString());
                temp.Merge(Type.Missing);
                temp.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                temp.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                temp.Value2 = "Официант: ";
                temp.Font.Name = "Calibri";
                temp.Font.Size = 8;

                temp = workSheet.get_Range("C" + r.ToString(), "H" + r.ToString());
                temp.Merge(Type.Missing);
                temp.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                temp.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                temp.Value2 = post_temp.GetValue(0).ToString() + " " + post_temp.GetValue(1).ToString();
                temp.Font.Name = "Calibri";
                temp.Font.Size = 8;

                r++;

                give_me_name = new OleDbCommand("SELECT [Фамилия], [Имя] FROM [Сотрудники] WHERE [id сотрудника]='" + post.GetValue(3).ToString() + "'", podkl);
                post_temp = give_me_name.ExecuteReader();
                post_temp.Read();

                temp = workSheet.get_Range("A" + r.ToString(), "B" + r.ToString());
                temp.Merge(Type.Missing);
                temp.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                temp.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                temp.Value2 = "Удалил: ";
                temp.Font.Name = "Calibri";
                temp.Font.Size = 8;

                temp = workSheet.get_Range("C" + r.ToString(), "H" + r.ToString());
                temp.Merge(Type.Missing);
                temp.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                temp.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                temp.Value2 = post_temp.GetValue(0).ToString() + " " + post_temp.GetValue(1).ToString();
                temp.Font.Name = "Calibri";
                temp.Font.Size = 8;

                r++;

                give_me_name = new OleDbCommand("SELECT [Причина] FROM [Причины удаления] WHERE [Код причины]='" + post.GetValue(4).ToString() + "'", podkl);
                post_temp = give_me_name.ExecuteReader();
                post_temp.Read();

                temp = workSheet.get_Range("A" + r.ToString(), "B" + r.ToString());
                temp.Merge(Type.Missing);
                temp.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                temp.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                temp.Value2 = "Причина: ";
                temp.Font.Name = "Calibri";
                temp.Font.Size = 8;

                temp = workSheet.get_Range("C" + r.ToString(), "H" + r.ToString());
                temp.Merge(Type.Missing);
                temp.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                temp.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                temp.Value2 = post_temp.GetValue(0).ToString();
                temp.Font.Name = "Calibri";
                temp.Font.Size = 8;

                r++;
                r++;
            }

            r++;
            //#############################################################
            temp = workSheet.get_Range("A" + r.ToString(), "B" + r.ToString());
            temp.Merge(Type.Missing);
            temp.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
            temp.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            temp.Value2 = "МЕНЕДЖЕР";
            temp.Font.Size = 8;
            temp.Borders.get_Item(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous;

            temp = workSheet.get_Range("C" + r.ToString(), "H" + r.ToString());
            temp.Merge(Type.Missing);
            temp.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
            temp.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            temp.Value2 = (Owner as select_stol).name;
            temp.Font.Size = 8;
            temp.Borders.get_Item(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous;

            for (int i = 1; i < 5; i++)
            {
                workSheet.Columns[i].ColumnWidth = 3.86;
            }
            for (int i = 5; i < 9; i++)
            {
                workSheet.Columns[i].ColumnWidth = 2.86;
            }

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

            podkl.Close();
            workBook.Close(false, Type.Missing, Type.Missing);

            excelApp.Quit();*/
            excelApp.Visible = true;
        }

        private void vyr_of_Click(object sender, EventArgs e)
        {
            if (tek_pol == 1)
            {
                print_all_vr();
                return;
            }
            tek_pol = 1;
            show_buttons();
            refresh();
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

        private void btn_Click(object sender, EventArgs e)
        {
            switch (tek_pol)
            {
                case (1):
                    var tt = vyr_oficianty[oficianty.FindIndex(t => t == (sender as Button).Text)];

                    podkl.Open();
                    komand = new OleDbCommand("SELECT [Фамилия], [Имя] FROM [Сотрудники] WHERE [id сотрудника]='" + (sender as Button).Text + "'", podkl);
                    post = komand.ExecuteReader();
                    post.Read();

                    string temp_name = post.GetValue(0).ToString() + " " + post.GetValue(1).ToString();

                    Excel.Application excelApp = new Excel.Application();
                    Excel.Workbook workBook = excelApp.Workbooks.Add(Type.Missing);
                    Excel.Worksheet workSheet = (Excel.Worksheet)excelApp.Worksheets.get_Item(1);

                    excelApp.Visible = false;

                    Excel.Range temp;
                    int r = 7;

                    temp = workSheet.get_Range("A1", "G2");
                    temp.Merge(Type.Missing);
                    temp.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    temp.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    temp.Value2 = "ООО «Воробушек»";
                    temp.Font.Name = "Calibri";
                    temp.Font.Size = 14;
                    temp.Font.Bold = true;
                    temp.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlDouble;

                    temp = workSheet.get_Range("A3", "G3");
                    temp.Merge(Type.Missing);
                    temp.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    temp.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    temp.Value2 = "Выручка по официантам";
                    temp.Font.Name = "Calibri";
                    temp.Font.Size = 9;

                    temp = workSheet.get_Range("A4", "G4");
                    temp.Merge(Type.Missing);
                    temp.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    temp.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    temp.Value2 = "Кассовый день " + DateTime.Now.ToShortDateString().ToString();
                    temp.Font.Name = "Calibri";
                    temp.Font.Size = 9;

                    temp = workSheet.get_Range("A5", "G5");
                    temp.Merge(Type.Missing);
                    temp.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    temp.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    temp.Value2 = DateTime.Now.ToString();
                    temp.Font.Name = "Calibri";
                    temp.Font.Size = 9;
                    temp.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlDash;

                    temp = workSheet.get_Range("A6", "E6");
                    temp.Merge(Type.Missing);
                    temp.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    temp.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    temp.Value2 = "Официант";
                    temp.Font.Name = "Calibri";
                    temp.Font.Size = 9;
                    temp.Font.Bold = true;
                    temp.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlDash;

                    temp = workSheet.get_Range("F6", "G6");
                    temp.Merge(Type.Missing);
                    temp.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    temp.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    temp.Value2 = "Сумма";
                    temp.Font.Name = "Calibri";
                    temp.Font.Size = 9;
                    temp.Font.Bold = true;
                    temp.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlDash;


                    temp = workSheet.get_Range("A" + r.ToString(), "E" + r.ToString());
                    temp.Merge(Type.Missing);
                    temp.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    temp.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    temp.Value2 = temp_name;
                    temp.Font.Name = "Calibri";
                    temp.Font.Size = 8;

                    temp = workSheet.get_Range("F" + r.ToString(), "G" + r.ToString());
                    temp.Merge(Type.Missing);
                    temp.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    temp.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    temp.Value2 = tt.ToString();
                    temp.Font.Name = "Calibri";
                    temp.Font.Size = 8;

                    r += 2;
                    //#############################################################
                    temp = workSheet.get_Range("A" + r.ToString(), "B" + r.ToString());
                    temp.Merge(Type.Missing);
                    temp.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                    temp.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    temp.Value2 = "МЕНЕДЖЕР";
                    temp.Font.Size = 8;
                    temp.Borders.get_Item(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous;

                    temp = workSheet.get_Range("C" + r.ToString(), "G" + r.ToString());
                    temp.Merge(Type.Missing);
                    temp.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                    temp.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    temp.Value2 = (Owner as select_stol).name;
                    temp.Font.Size = 8;
                    temp.Borders.get_Item(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous;

                    podkl.Close();

                    for (int i = 1; i <= 8; i++)
                    {
                        workSheet.Columns[i].ColumnWidth = 3.86;
                    }

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

                    break;
            }
        }



        private void vyr_all_Click(object sender, EventArgs e)
        {
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook workBook = excelApp.Workbooks.Add(Type.Missing);
            Excel.Worksheet workSheet = (Excel.Worksheet)excelApp.Worksheets.get_Item(1);

            excelApp.Visible = false;

            Excel.Range temp;
            int r = 7;

            temp = workSheet.get_Range("A1", "G2");
            temp.Merge(Type.Missing);
            temp.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            temp.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            temp.Value2 = "ООО «Воробушек»";
            temp.Font.Name = "Calibri";
            temp.Font.Size = 14;
            temp.Font.Bold = true;
            temp.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlDouble;

            temp = workSheet.get_Range("A3", "G3");
            temp.Merge(Type.Missing);
            temp.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            temp.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            temp.Value2 = "Общая выручка";
            temp.Font.Name = "Calibri";
            temp.Font.Size = 9;

            temp = workSheet.get_Range("A4", "G4");
            temp.Merge(Type.Missing);
            temp.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            temp.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            temp.Value2 = "Кассовый день " + DateTime.Now.ToShortDateString().ToString();
            temp.Font.Name = "Calibri";
            temp.Font.Size = 9;

            temp = workSheet.get_Range("A5", "G5");
            temp.Merge(Type.Missing);
            temp.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            temp.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            temp.Value2 = DateTime.Now.ToString();
            temp.Font.Name = "Calibri";
            temp.Font.Size = 9;
            temp.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlDash;

            temp = workSheet.get_Range("A6", "E6");
            temp.Merge(Type.Missing);
            temp.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            temp.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            temp.Value2 = "Тип";
            temp.Font.Name = "Calibri";
            temp.Font.Size = 9;
            temp.Font.Bold = true;
            temp.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlDash;

            temp = workSheet.get_Range("F6", "G6");
            temp.Merge(Type.Missing);
            temp.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            temp.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            temp.Value2 = "Сумма";
            temp.Font.Name = "Calibri";
            temp.Font.Size = 9;
            temp.Font.Bold = true;
            temp.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlDash;

            podkl.Open();
            komand = new OleDbCommand("SELECT [Заказы].[Оплата картой], Sum([Заказы].[Сумма]) AS [Sum-Сумма] FROM [Заказы] GROUP BY [Заказы].[Оплата картой]", podkl);
            post = komand.ExecuteReader();
            while (post.Read())
            {
                temp = workSheet.get_Range("A" + r.ToString(), "E" + r.ToString());
                temp.Merge(Type.Missing);
                temp.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                temp.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                if (post.GetValue(0).ToString().ToLower() == "true")
                    temp.Value2 = "Карта";
                else
                    temp.Value2 = "Наличные";
                temp.Font.Name = "Calibri";
                temp.Font.Size = 8;

                temp = workSheet.get_Range("F" + r.ToString(), "G" + r.ToString());
                temp.Merge(Type.Missing);
                temp.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                temp.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                if (post.GetValue(1).ToString() == "")
                    temp.Value2 = 0;
                else
                    temp.Value2 = int.Parse(post.GetValue(1).ToString());
                temp.Font.Name = "Calibri";
                temp.Font.Size = 8;
                r++;
            }

            podkl.Close();
            r += 1;
            //#############################################################
            temp = workSheet.get_Range("A" + r.ToString(), "B" + r.ToString());
            temp.Merge(Type.Missing);
            temp.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
            temp.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            temp.Value2 = "МЕНЕДЖЕР";
            temp.Font.Size = 8;
            temp.Borders.get_Item(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous;

            temp = workSheet.get_Range("C" + r.ToString(), "G" + r.ToString());
            temp.Merge(Type.Missing);
            temp.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
            temp.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            temp.Value2 = (Owner as select_stol).name;
            temp.Font.Size = 8;
            temp.Borders.get_Item(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous;

            for (int i = 1; i < 8; i++)
            {
                workSheet.Columns[i].ColumnWidth = 3.86;
            }

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








        private void rashod_Click(object sender, EventArgs e)
        {
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook workBook = excelApp.Workbooks.Add(Type.Missing);
            Excel.Worksheet workSheet = (Excel.Worksheet)excelApp.Worksheets.get_Item(1);

            excelApp.Visible = false;

            Excel.Range temp;
            int r = 7;

            temp = workSheet.get_Range("A1", "H2");
            temp.Merge(Type.Missing);
            temp.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            temp.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            temp.Value2 = "ООО «Воробушек»";
            temp.Font.Name = "Calibri";
            temp.Font.Size = 14;
            temp.Font.Bold = true;
            temp.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlDouble;

            temp = workSheet.get_Range("A3", "H3");
            temp.Merge(Type.Missing);
            temp.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            temp.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            temp.Value2 = "Расход блюд";
            temp.Font.Name = "Calibri";
            temp.Font.Size = 9;

            temp = workSheet.get_Range("A4", "H4");
            temp.Merge(Type.Missing);
            temp.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            temp.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            temp.Value2 = "Кассовый день " + DateTime.Now.ToShortDateString().ToString();
            temp.Font.Name = "Calibri";
            temp.Font.Size = 9;

            temp = workSheet.get_Range("A5", "H5");
            temp.Merge(Type.Missing);
            temp.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            temp.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            temp.Value2 = DateTime.Now.ToString();
            temp.Font.Name = "Calibri";
            temp.Font.Size = 9;
            temp.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlDash;

            temp = workSheet.get_Range("A6", "D6");
            temp.Merge(Type.Missing);
            temp.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            temp.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            temp.Value2 = "Блюдо";
            temp.Font.Name = "Calibri";
            temp.Font.Size = 9;
            temp.Font.Bold = true;
            temp.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlDash;

            temp = workSheet.get_Range("E6", "F6");
            temp.Merge(Type.Missing);
            temp.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            temp.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            temp.Value2 = "Кол-во";
            temp.Font.Name = "Calibri";
            temp.Font.Size = 9;
            temp.Font.Bold = true;
            temp.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlDash;

            temp = workSheet.get_Range("G6", "H6");
            temp.Merge(Type.Missing);
            temp.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            temp.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            temp.Value2 = "Сумма";
            temp.Font.Name = "Calibri";
            temp.Font.Size = 9;
            temp.Font.Bold = true;
            temp.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlDash;

            podkl.Open();
            komand = new OleDbCommand("SELECT [Блюда].[Наименование], Sum([Состав заказа].[Количество]) AS [Sum-Количество], Sum(Блюда.Цена) AS [Sum-Цена] FROM [Блюда] INNER JOIN [Состав заказа] ON [Блюда].[Код блюда] = [Состав заказа].[Код блюда] GROUP BY [Состав заказа].[Код блюда], [Блюда].[Наименование]", podkl);
            post = komand.ExecuteReader();
            while (post.Read())
            {
                temp = workSheet.get_Range("A" + r.ToString(), "D" + r.ToString());
                temp.Merge(Type.Missing);
                temp.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                temp.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                temp.Value2 = post.GetValue(0).ToString();
                temp.Font.Name = "Calibri";
                temp.Font.Size = 7;

                temp = workSheet.get_Range("E" + r.ToString(), "F" + r.ToString());
                temp.Merge(Type.Missing);
                temp.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                temp.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                temp.Value2 = post.GetValue(1).ToString(); ;
                temp.Font.Name = "Calibri";
                temp.Font.Size = 7;

                temp = workSheet.get_Range("G" + r.ToString(), "H" + r.ToString());
                temp.Merge(Type.Missing);
                temp.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                temp.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                temp.Value2 = post.GetValue(2).ToString(); ;
                temp.Font.Name = "Calibri";
                temp.Font.Size = 7;
                r++;
            }

            r++;
            //#############################################################
            temp = workSheet.get_Range("A" + r.ToString(), "B" + r.ToString());
            temp.Merge(Type.Missing);
            temp.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
            temp.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            temp.Value2 = "МЕНЕДЖЕР";
            temp.Font.Size = 8;
            temp.Borders.get_Item(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous;

            temp = workSheet.get_Range("C" + r.ToString(), "H" + r.ToString());
            temp.Merge(Type.Missing);
            temp.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
            temp.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            temp.Value2 = (Owner as select_stol).name;
            temp.Font.Size = 8;
            temp.Borders.get_Item(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous;

            for (int i = 1; i < 5; i++)
            {
                workSheet.Columns[i].ColumnWidth = 3.86;
            }
            for (int i = 5; i < 9; i++)
            {
                workSheet.Columns[i].ColumnWidth = 2.86;
            }

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

            podkl.Close();
        }
        private void print_all_vr()
        {
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook workBook = excelApp.Workbooks.Add(Type.Missing);
            Excel.Worksheet workSheet = (Excel.Worksheet)excelApp.Worksheets.get_Item(1);

            excelApp.Visible = false;

            Excel.Range temp;
            int r = 7;

            temp = workSheet.get_Range("A1", "G2");
            temp.Merge(Type.Missing);
            temp.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            temp.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            temp.Value2 = "ООО «Воробушек»";
            temp.Font.Name = "Calibri";
            temp.Font.Size = 14;
            temp.Font.Bold = true;
            temp.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlDouble;

            temp = workSheet.get_Range("A3", "G3");
            temp.Merge(Type.Missing);
            temp.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            temp.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            temp.Value2 = "Выручка по официантам";
            temp.Font.Name = "Calibri";
            temp.Font.Size = 9;

            temp = workSheet.get_Range("A4", "G4");
            temp.Merge(Type.Missing);
            temp.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            temp.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            temp.Value2 = "Кассовый день " + DateTime.Now.ToShortDateString().ToString();
            temp.Font.Name = "Calibri";
            temp.Font.Size = 9;

            temp = workSheet.get_Range("A5", "G5");
            temp.Merge(Type.Missing);
            temp.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            temp.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            temp.Value2 = DateTime.Now.ToString();
            temp.Font.Name = "Calibri";
            temp.Font.Size = 9;
            temp.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlDash;

            temp = workSheet.get_Range("A6", "E6");
            temp.Merge(Type.Missing);
            temp.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            temp.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            temp.Value2 = "Официант";
            temp.Font.Name = "Calibri";
            temp.Font.Size = 9;
            temp.Font.Bold = true;
            temp.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlDash;

            temp = workSheet.get_Range("F6", "G6");
            temp.Merge(Type.Missing);
            temp.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            temp.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            temp.Value2 = "Сумма";
            temp.Font.Name = "Calibri";
            temp.Font.Size = 9;
            temp.Font.Bold = true;
            temp.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlDash;

            for (int i = 0; i < oficianty.Count; i++)
            {
                podkl.Open();
                komand = new OleDbCommand("SELECT [Фамилия], [Имя] FROM [Сотрудники] WHERE [id сотрудника]='" + oficianty[i] + "'", podkl);
                post = komand.ExecuteReader();
                post.Read();
                string temp_name = post.GetValue(0).ToString() + " " + post.GetValue(1).ToString();

                var tt = vyr_oficianty[i];


                temp = workSheet.get_Range("A" + r.ToString(), "E" + r.ToString());
                temp.Merge(Type.Missing);
                temp.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                temp.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                temp.Value2 = temp_name;
                temp.Font.Name = "Calibri";
                temp.Font.Size = 8;

                temp = workSheet.get_Range("F" + r.ToString(), "G" + r.ToString());
                temp.Merge(Type.Missing);
                temp.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                temp.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                temp.Value2 = tt.ToString();
                temp.Font.Name = "Calibri";
                temp.Font.Size = 8;
                r++;
                podkl.Close();
            }
            r += 1;
            //#############################################################
            temp = workSheet.get_Range("A" + r.ToString(), "B" + r.ToString());
            temp.Merge(Type.Missing);
            temp.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
            temp.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            temp.Value2 = "МЕНЕДЖЕР";
            temp.Font.Size = 8;
            temp.Borders.get_Item(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous;

            temp = workSheet.get_Range("C" + r.ToString(), "G" + r.ToString());
            temp.Merge(Type.Missing);
            temp.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
            temp.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            temp.Value2 = (Owner as select_stol).name;
            temp.Font.Size = 8;
            temp.Borders.get_Item(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous;

            for (int i = 1; i < 8; i++)
            {
                workSheet.Columns[i].ColumnWidth = 3.86;
            }

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
    }
}
