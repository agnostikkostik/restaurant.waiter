using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;

namespace WindowsFormsApplication1
{
    public class kategoria_1
    {
        public string name;
        public int kod;
        public List<podcategroia_1> podcategorii = new List<podcategroia_1>();
        public kategoria_1(string temp_name, int temp_kod)
        {
            name = temp_name;
            kod = temp_kod;
        }
    }

    public class podcategroia_1
    {
        public string name;
        public int kod;
        public List<blydo_1> blyda = new List<blydo_1>();

        public podcategroia_1(string temp_name, int temp_kod)
        {
            name = temp_name;
            kod = temp_kod;
        }
    }

    public class blydo_1
    {
        public int kod;
        public string name;
        public string name_in_chek;
        public int cena;
        public string[][] modifikatory_ob;
        public List<blydo_1>[] modifikatory_ob_List;
        public string[][] modifikatory_nob;
        public List<blydo_1>[] modifikatory_nob_List;

        public blydo_1(int temp_kod, string temp_name, string temp_name_in_chek, int temp_cena, string temp_mod_ob, string temp_mod_nob, OleDbConnection temp_podkl)
        {
            kod = temp_kod;
            name = temp_name;
            name_in_chek = temp_name_in_chek;
            cena = temp_cena;

            #region обязательные модификаторы
            if (temp_mod_ob == "")
                modifikatory_ob = null;
            else
            {
                bool other_bl = false;
                for (int i = 0; i <= 9; i++)
                {
                    if (temp_mod_ob.IndexOf(Char.Parse(i.ToString())) != -1)
                        other_bl = true;
                }

                #region МОДИФИКАТОРЫ-БЛЮДА
                if (other_bl)
                {
                    modifikatory_ob = null;
                    modifikatory_nob = null;

                    string[] temp = temp_mod_ob.Split(';');
                    modifikatory_ob_List = new List<blydo_1>[temp.Length];

                    for (int i = 0; i < temp.Length; i++)
                    {
                        string[] temp2 = temp[i].Split(',');
                        modifikatory_ob_List[i] = new List<blydo_1>();

                        for (int j = 0; j < temp2.Length; j++)
                        {

                            OleDbCommand komand;
                            OleDbDataReader post;
                            OleDbConnection podkl = temp_podkl;

                            komand = new OleDbCommand("Select [Код блюда], [Наименование], [Наименование в чеке], [Цена], [Обязательные модификаторы], [Необязательные модификаторы] From [Блюда] WHERE [Код блюда]=" + temp2[j].ToString(), podkl);
                            post = komand.ExecuteReader();
                            while (post.Read())
                            {
                                int kod_2 = int.Parse(post.GetValue(0).ToString());
                                string naimenovanie_2 = post.GetValue(1).ToString();
                                string naimenovanie_chek_2 = post.GetValue(2).ToString();
                                int cena_2 = int.Parse(post.GetValue(3).ToString());
                                string modifikator_ob_2 = post.GetValue(4).ToString();
                                string modifikator_nob_2 = post.GetValue(5).ToString();

                                modifikatory_ob_List[i].Add(new blydo_1(kod_2, naimenovanie_2, naimenovanie_chek_2, cena_2, modifikator_ob_2, modifikator_nob_2, podkl));
                            }
                            post.Close();
                        }
                    }
                }
                #endregion
                #region ОБЫЧНЫЕ МОДИФИКАТОРЫ
                else
                {
                    string[] temp = temp_mod_ob.Split(';');
                    modifikatory_ob = new string[temp.Length][];

                    for (int i = 0; i < temp.Length; i++)
                    {
                        string[] temp2 = temp[i].Split(',');
                        modifikatory_ob[i] = new string[temp2.Length];
                        modifikatory_ob[i] = temp2;
                    }
                }
                #endregion
            }
            #endregion

            #region необязательные модификаторы
            if (temp_mod_nob == "")
                modifikatory_nob = null;
            else
            {
                string[] temp = temp_mod_nob.Split(';');
                modifikatory_nob = new string[temp.Length][];

                for (int i = 0; i < temp.Length; i++)
                {
                    string[] temp2 = temp[i].Split(',');
                    modifikatory_nob[i] = new string[temp2.Length];
                    modifikatory_nob[i] = temp2;
                }
            }
            #endregion
        }
    }
}
