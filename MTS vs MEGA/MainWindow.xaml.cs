using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;

namespace MTS_vs_MEGA
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        string[] R1C1 = new string[] { "0", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN", "AO", "AP", "AQ", "AR", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "AZ", "BA", "BB", "BC", "BD", "BE", "BF", "BG", "BH", "BI", "BJ", "BK", "BL", "BM", "BN", "BO", "BP", "BQ", "BR", "BS", "BT", "BU", "BV", "BW", "BX", "BY", "BZ", "CA", "CB", "CC", "CD", "CE", "CF", "CG", "CH", "CI", "CJ", "CK", "CL", "CM", "CN", "CO", "CP", "CQ", "CR", "CS", "CT", "CU", "CV", "CW", "CX", "CY", "CZ", "DA", "DB", "DC", "DD", "DE", "DF", "DG", "DH", "DI", "DJ", "DK", "DL", "DM", "DN", "DO", "DP", "DQ", "DR", "DS", "DT", "DU", "DV", "DW", "DX", "DY", "DZ", "EA", "EB", "EC", "ED", "EE", "EF", "EG", "EH", "EI", "EJ", "EK", "EL", "EM", "EN", "EO", "EP", "EQ", "ER", "ES", "ET", "EU", "EV", "EW", "EX", "EY", "EZ" };


        public MainWindow()
        {
            InitializeComponent();
        }

        public class condition
        {
            public bool upcond;
            public bool downcond;
            public string medText;
            public string downText;
            public int num;

            public condition(bool UPcond, bool DOWNcond, string medTEXT, string downTEXT)
            {
                upcond = UPcond;
                downcond = DOWNcond;
                medText = medTEXT;
                downText = downTEXT;

                if (UPcond)
                    num = 1;
                else if (DOWNcond)
                    num = 2;
                else
                    num = 3;
            }
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            //MessageBox.Show((Convert.ToDouble(procKach2.Text) / Convert.ToDouble(100)).ToString());

            MessageBox.Show("Укажите результат МТС");
            string MTSfile = new OpenExcelFile().Filenamereturn();
            if (MTSfile == "can not open file")
                return;
            object[][] MTS = getarray(MTSfile, 1, new int[] { 1,3,4,15,17,21 });

            //MessageBox.Show(MTS[4][0].ToString());

            MessageBox.Show("Укажите результат МЕГАФОН");
            string MEGAfile = new OpenExcelFile().Filenamereturn();
            if (MEGAfile == "can not open file")
                return;
            object[][] MEGA = getarray(MEGAfile, 1, new int[] { 1, 3, 4, 15, 17, 21 });

            MessageBox.Show("Укажите файл анализа");
            string Spisok = new OpenExcelFile().Filenamereturn();
            if (Spisok == "can not open file")
                return;
            object[][] spisok = getarray(Spisok, 1, new int[] { 1 });

            string[,] NEW = new string[spisok[0].Length,4];

            int c = 0;

            foreach (object sp in spisok[0])
            {
                bool mtsfind = false;
                for (int ro = 0; ro < MTS[0].Length; ro ++)
                {
                    if (sp.ToString() == MTS[0][ro].ToString())
                    {
                        List<condition> conditionsMTS = new List<condition>();

                        conditionsMTS.Add(new condition(Convert.ToDouble(MTS[4][ro]) >= Convert.ToDouble(podMTS1.Text),
                            Convert.ToDouble(MTS[4][ro]) > Convert.ToDouble(podMTS2.Text),
                            "Среднее подобие МТС, ", "Плохое подобие МТС, "));

                        conditionsMTS.Add(new condition(Convert.ToDouble(MTS[5][ro]) >= Convert.ToDouble(sredPop1.Text),
                            Convert.ToDouble(MTS[5][ro]) > Convert.ToDouble(sredPop2.Text),
                            "Среднее пополнение МТС, ", "Плохое пополнение МТС, "));

                        conditionsMTS.Add(new condition((Convert.ToDouble(MTS[1][ro]) / Convert.ToDouble(MTS[2][ro])) >= (Convert.ToDouble(procKach1.Text) / 100),
                           (Convert.ToDouble(MTS[1][ro]) / Convert.ToDouble(MTS[2][ro])) > (Convert.ToDouble(procKach2.Text) / Convert.ToDouble(100)),
                            "Средний процент качества МТС, ", "Плохой процент качества МТС, "));

                        conditionsMTS.Add(new condition(DoubleFromProc(MTS[3][ro].ToString()) >= Convert.ToDouble(m2act1.Text),
                            DoubleFromProc(MTS[3][ro].ToString()) > Convert.ToDouble(m2act2.Text),
                            "Средня 2М активность МТС, ", "Плохая 2М активность МТС, "));


                        bool find = false;
                        for (int i = 0; i < MEGA[0].Length; i++)
                        {
                            if (MTS[0][ro].ToString() == MEGA[0][i].ToString())
                            {
                                List<condition> conditionsMEGA = new List<condition>();

                                conditionsMEGA.Add(new condition(Convert.ToDouble(MEGA[4][i]) >= Convert.ToDouble(podMEGA1.Text),
                                    Convert.ToDouble(MEGA[4][i]) > Convert.ToDouble(podMEGA2.Text),
                                    "Среднее подобие МТС, ", "Плохое подобие МТС, "));

                                conditionsMEGA.Add(new condition(Convert.ToDouble(MEGA[5][i]) >= Convert.ToDouble(sredPop1.Text),
                                    Convert.ToDouble(MEGA[5][i]) > Convert.ToDouble(sredPop2.Text),
                                    "Среднее пополнение МТС, ", "Плохое пополнение МТС, "));

                                conditionsMEGA.Add(new condition((Convert.ToDouble(MEGA[1][i]) / Convert.ToDouble(MEGA[2][i])) >= (Convert.ToDouble(procKach1.Text) / 100),
                                   (Convert.ToDouble(MEGA[1][i]) / Convert.ToDouble(MEGA[2][i])) > (Convert.ToDouble(procKach2.Text) / 100),
                                    "Средний процент качества МТС, ", "Плохой процент качества МТС, "));

                                conditionsMEGA.Add(new condition(DoubleFromProc(MEGA[3][i].ToString()) >= Convert.ToDouble(m2act1.Text),
                                    DoubleFromProc(MEGA[3][i].ToString()) > Convert.ToDouble(m2act2.Text),
                                    "Средня 2М активность, ", "Плохая 2М активность, "));


                                if (conditionsMTS[0].upcond && conditionsMTS[1].upcond && conditionsMTS[2].upcond && conditionsMTS[3].upcond &&
                                    conditionsMEGA[0].upcond && conditionsMEGA[1].upcond && conditionsMEGA[2].upcond && conditionsMEGA[3].upcond)
                                {
                                    NEW[c, 1] = "green";
                                }
                                else if (!conditionsMTS[0].downcond || !conditionsMTS[1].downcond || !conditionsMTS[2].downcond || !conditionsMTS[3].downcond ||
                                    !conditionsMEGA[0].downcond || !conditionsMEGA[1].downcond || !conditionsMEGA[2].downcond || !conditionsMEGA[3].downcond)
                                {
                                    NEW[c, 1] = "red";
                                    foreach (condition con in conditionsMTS)
                                    {
                                        NEW[c, 2] += (!con.downcond) ? con.downText : "";
                                    }
                                    foreach (condition con in conditionsMEGA)
                                    {
                                        NEW[c, 2] += (!con.downcond) ? con.downText : "";
                                    }
                                    foreach (condition con in conditionsMTS)
                                    {
                                        NEW[c, 2] += (!con.upcond && con.downcond) ? con.medText : "";
                                    }
                                    foreach (condition con in conditionsMEGA)
                                    {
                                        NEW[c, 2] += (!con.upcond && con.downcond) ? con.medText : "";
                                    }
                                }
                                else
                                {
                                    NEW[c, 1] = "yellow";
                                    foreach (condition con in conditionsMTS)
                                    {
                                        NEW[c, 2] += (!con.upcond && con.downcond) ? con.medText : "";
                                    }
                                    foreach (condition con in conditionsMEGA)
                                    {
                                        NEW[c, 2] += (!con.upcond && con.downcond) ? con.medText : "";
                                    }
                                }

                                find = true;
                                break;
                            }
                        }


                        if (!find)
                        {
                            NEW[4, 1] += "Продает только МТС!, ";
                            if (conditionsMTS[0].upcond && conditionsMTS[1].upcond && conditionsMTS[2].upcond && conditionsMTS[3].upcond)
                            {
                                NEW[c, 1] = "green mts only";
                            }
                            else if (!conditionsMTS[0].downcond || !conditionsMTS[1].downcond || !conditionsMTS[2].downcond || !conditionsMTS[3].downcond)
                            {
                                NEW[c, 1] = "red mts only";
                                foreach (condition con in conditionsMTS)
                                {
                                    NEW[c, 2] += (!con.downcond) ? con.downText : "";
                                }
                                foreach (condition con in conditionsMTS)
                                {
                                    NEW[c, 2] += (!con.upcond && con.downcond) ? con.medText : "";
                                }
                            }
                            else
                            {
                                NEW[c, 1] = "yellow mts only";
                                foreach (condition con in conditionsMTS)
                                {
                                    NEW[c, 2] += (!con.upcond && con.downcond) ? con.medText : "";
                                }
                            }
                        }
                        mtsfind = true;
                        break;
                    }
                }
                c++;
            }

            MessageBox.Show("Укажите файл для вставки");

            Spisok = new OpenExcelFile().Filenamereturn();
            if (Spisok == "can not open file")
                return;

            insert(Spisok, NEW, c, 2);

            MessageBox.Show("Конец программы");


        }


        public void insert(string path, object[,] arr, int rows, int col)
        {
            #region Открытие Excel
            var ExcelApp = new Excel.Application();
            ExcelApp.Visible = false;
            Excel.Sheets excelsheets;
            Excel.Worksheet excelworksheet;
            //Excel.Workbooks workbooks;
            Excel.Workbook book;
            Excel.Range range = null;

            book = ExcelApp.Workbooks.Open(path, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            //book.ActiveSheet.get_Item(1);
            excelsheets = book.Worksheets;
            excelworksheet = (Excel.Worksheet)excelsheets.get_Item(1);

            #endregion
            Process[] List = Process.GetProcessesByName("EXCEL");

            int Columns = excelworksheet.UsedRange.Columns.Count;

            /*
            range = null;
            range = excelworksheet.get_Range(R1C1[1] + "1:" + R1C1[col] + "1");
            range.Value2 = new object[,] { {"Дилер Дистр", "Всего платежей" ,"Хорошие симки за период", "Всего симок в комиссии","кол-во симок >120р в первом месяце" ,"кол-во симок >120р во втором месяце","кол-во симок >120р в третьем месяце",
             "кол-во симок >120р в 4-6 месяце"  ,"кол-во симок >120р в 7-12 месяце", "6) платежи на комис" ,"7) платежи на отгрузки" ,
                    "8) хорошие (>120р) симки 1-го пер набл на кол-во отгрузок" ,"9) хорошие (>120р) симки 1,2,3 пер набл на кол-во отгрузок","1м активность","2м активность","3м активность","Подобие","2м активность","тарифы с АП",
                    "тарифы без АП","Среднее пополнение","Тариф 1","Тариф 2","1м:2м:3м:4+м (ком)","в комиссии | отгрузки за период" } };
            */
            range = null;
            range = excelworksheet.get_Range(R1C1[Columns+1] + "2:" + R1C1[Columns+col+1] + rows.ToString());
            range.Value2 = arr;


            #region Закрытие Excel

            book.Save();
            book.Close(false, false, false);

            ExcelApp.Quit();

            ExcelApp = null;
            excelsheets = null;
            excelworksheet = null;
            //workbooks = null;
            book = null;
            range = null;
            #endregion
            CloseProcess(List);
        }


        private object[][] getarray(string path, int list, int[] columns) //возвращает массив указаных колонок 
        {
            #region Открытие Excel
            var ExcelApp = new Excel.Application();
            ExcelApp.Visible = false;
            Excel.Sheets excelsheets;
            Excel.Worksheet excelworksheet;
            //Excel.Workbooks workbooks;
            Excel.Workbook book;
            Excel.Range range = null;

            book = ExcelApp.Workbooks.Open(path, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            //book.ActiveSheet.get_Item(list);
            excelsheets = book.Worksheets;
            excelworksheet = (Excel.Worksheet)excelsheets.get_Item(list);

            #endregion
            Process[] List = Process.GetProcessesByName("EXCEL");

            int Rows = excelworksheet.UsedRange.Rows.Count;
            int Columns = excelworksheet.UsedRange.Columns.Count;
            object[][] arr = new object[columns.Length][];

            int icolumn = 0;
            foreach (int column in columns)
            {
                for (int i = 0; i < Columns + 1; i++)
                {
                    if (column == i)
                    {
                        object[,] massiv;
                        arr[icolumn] = new object[Rows];
                        range = excelworksheet.get_Range(R1C1[i] + "2:" + R1C1[i] + Rows.ToString());
                        massiv = (System.Object[,])range.get_Value(Type.Missing);
                        arr[icolumn] = massiv.Cast<object>().ToArray();
                        icolumn++;
                    }
                }
            }

            #region Закрытие Excel

            book.Close(false, false, false);

            ExcelApp.Quit();


            ExcelApp = null;
            excelsheets = null;
            excelworksheet = null;
            book = null;
            range = null;
            #endregion
            CloseProcess(List);

            return arr;
        }



        public void CloseProcess(Process[] before) //закрытие массива процессов (для закрытия процессов EXCEL) 
        {
            Process[] List;
            List = Process.GetProcessesByName("EXCEL");
            foreach (Process proc in List)
            {
                if (!before.Contains(proc))
                    proc.Kill();
            }
        }

        public double DoubleFromProc(string s)
        {
            s = s.Substring(0, s.IndexOf("%"));
            return Convert.ToDouble(s);
        }

    }
}
