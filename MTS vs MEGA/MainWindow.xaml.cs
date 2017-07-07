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
            //MessageBox.Show(DoubleFromProc("20% (88)").ToString());


            string MTSfile = new OpenExcelFile().Filenamereturn();
            if (MTSfile == "can not open file")
                return;
            object[][] MTS = getarray(MTSfile, 1, new int[] { 3, 1 });

            string MEGAfile = new OpenExcelFile().Filenamereturn();
            if (MEGAfile == "can not open file")
                return;
            object[][] MEGA = getarray(MTSfile, 1, new int[] { 3, 1 });


            string[,] NEW = new string[111,4];

            int c = 0;
            foreach (string name in MTS[0])
            {
                List<condition> conditionsMTS = new List<condition>();

                conditionsMTS.Add(new condition(Convert.ToDouble(MTS[c][16]) >= Convert.ToDouble(podMTS1), 
                    Convert.ToDouble(MTS[c][16]) > Convert.ToDouble(podMTS2), 
                    "Среднее подобие МТС, ", "Плохое подобие МТС, "));

                conditionsMTS.Add(new condition(Convert.ToDouble(MTS[c][20]) >= Convert.ToDouble(sredPop1),
                    Convert.ToDouble(MTS[c][20]) > Convert.ToDouble(sredPop2),
                    "Среднее пополнение МТС, ", "Плохое пополнение МТС, "));

                conditionsMTS.Add(new condition((Convert.ToDouble(MTS[c][2])/ Convert.ToDouble(MTS[c][3])) >= (Convert.ToDouble(procKach1) / 100),
                   (Convert.ToDouble(MTS[c][2]) / Convert.ToDouble(MTS[c][3])) > (Convert.ToDouble(procKach2) / 100),
                    "Средний процент качества МТС, ", "Плохой процент качества МТС, "));

                conditionsMTS.Add(new condition(DoubleFromProc(MTS[c][14].ToString()) >= Convert.ToDouble(m2act1),
                    DoubleFromProc(MTS[c][14].ToString()) > Convert.ToDouble(m2act2),
                    "Средня 2М активность МТС, ", "Плохая 2М активность МТС, "));


                bool find = false;
                for (int i = 0; i < MEGA[0].Length; i++)
                {
                    if (name == MEGA[0][i].ToString())
                    {
                        List<condition> conditionsMEGA = new List<condition>();

                        conditionsMEGA.Add(new condition(Convert.ToDouble(MEGA[c][16]) >= Convert.ToDouble(podMEGA1),
                            Convert.ToDouble(MEGA[c][16]) > Convert.ToDouble(podMEGA2),
                            "Среднее подобие МТС, ", "Плохое подобие МТС, "));

                        conditionsMEGA.Add(new condition(Convert.ToDouble(MEGA[c][20]) >= Convert.ToDouble(sredPop1),
                            Convert.ToDouble(MEGA[c][20]) > Convert.ToDouble(sredPop2),
                            "Среднее пополнение МТС, ", "Плохое пополнение МТС, "));

                        conditionsMEGA.Add(new condition((Convert.ToDouble(MEGA[c][2]) / Convert.ToDouble(MEGA[c][3])) >= (Convert.ToDouble(procKach1) / 100),
                           (Convert.ToDouble(MEGA[c][2]) / Convert.ToDouble(MEGA[c][3])) > (Convert.ToDouble(procKach2) / 100),
                            "Средний процент качества МТС, ", "Плохой процент качества МТС, "));

                        conditionsMEGA.Add(new condition(DoubleFromProc(MEGA[c][14].ToString()) >= Convert.ToDouble(m2act1),
                            DoubleFromProc(MEGA[c][14].ToString()) > Convert.ToDouble(m2act2),
                            "Средня 2М активность, ", "Плохая 2М активность, "));


                        if (conditionsMTS[0].upcond && conditionsMTS[1].upcond && conditionsMTS[2].upcond && conditionsMTS[3].upcond &&
                            conditionsMEGA[0].upcond && conditionsMEGA[1].upcond && conditionsMEGA[2].upcond && conditionsMEGA[3].upcond)
                        {
                            NEW[c, 2] = "green";
                        }
                        else if (!conditionsMTS[0].downcond || !conditionsMTS[1].downcond || !conditionsMTS[2].downcond || !conditionsMTS[3].downcond ||
                            !conditionsMEGA[0].downcond || !conditionsMEGA[1].downcond || !conditionsMEGA[2].downcond || !conditionsMEGA[3].downcond)
                        {
                            NEW[c, 2] = "red";
                            foreach(condition con in conditionsMTS)
                            {
                                NEW[c, 4] += (!con.downcond) ? con.downText : "";
                            }
                            foreach (condition con in conditionsMEGA)
                            {
                                NEW[c, 4] += (!con.downcond) ? con.downText : "";
                            }
                            foreach (condition con in conditionsMTS)
                            {
                                NEW[c, 4] += (!con.upcond && con.downcond) ? con.medText : "";
                            }
                            foreach (condition con in conditionsMEGA)
                            {
                                NEW[c, 4] += (!con.upcond && con.downcond) ? con.medText : "";
                            }
                        }
                        else
                        {
                            NEW[c, 2] = "yellow";
                            foreach (condition con in conditionsMTS)
                            {
                                NEW[c, 4] += (!con.upcond && con.downcond) ? con.medText : "";
                            }
                            foreach (condition con in conditionsMEGA)
                            {
                                NEW[c, 4] += (!con.upcond && con.downcond) ? con.medText : "";
                            }
                        }

                        find = true;
                        break;
                    }
                }


                if (!find)
                {
                    NEW[4, c] += "Продает только МТС!, ";
                    if (conditionsMTS[0].upcond && conditionsMTS[1].upcond && conditionsMTS[2].upcond && conditionsMTS[3].upcond)
                    {
                        NEW[c, 2] = "green mts only";
                    }
                    else if (!conditionsMTS[0].downcond || !conditionsMTS[1].downcond || !conditionsMTS[2].downcond || !conditionsMTS[3].downcond)
                    {
                        NEW[c, 2] = "red mts only";
                        foreach (condition con in conditionsMTS)
                        {
                            NEW[c, 4] += (!con.downcond) ? con.downText : "";
                        }
                        foreach (condition con in conditionsMTS)
                        {
                            NEW[c, 4] += (!con.upcond && con.downcond) ? con.medText : "";
                        }
                    }
                    else
                    {
                        NEW[c, 2] = "yellow mts only";
                        foreach (condition con in conditionsMTS)
                        {
                            NEW[c, 4] += (!con.upcond && con.downcond) ? con.medText : "";
                        }
                    }
                }
                c++;
            }


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
            return Convert.ToDouble(s) / 100;
        }

    }
}
