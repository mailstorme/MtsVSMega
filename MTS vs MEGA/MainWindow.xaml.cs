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

namespace MTS_vs_MEGA
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
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

            public condition(bool UPcond, bool DOWNcond, string medTEXT, string downTEXT)
            {
                upcond = UPcond;
                downcond = DOWNcond;
                medText = medTEXT;
                downText = downTEXT;
            }
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            string[][] MTS;
            string[][] MEGA;
            string[,] NEW = new string[111,4];

            int c = 0;
            foreach (string name in MTS[0])
            {
                bool find = false;
                for (int i = 0; i < MEGA[0].Length; i++)
                {
                    bool[] Amts = new bool[6];
                    int l = 0;
                    Amts[l++] = Convert.ToDouble(MTS[c][2]) > Convert.ToDouble(podMTS1);
                    Amts[l++] = Convert.ToDouble(MTS[c][12]) > Convert.ToDouble(sredPop1);
                    Amts[l++] = Convert.ToDouble(MTS[c][3]) > Convert.ToDouble(m2act1);
                    Amts[l++] = Convert.ToDouble(MTS[c][111]) > Convert.ToDouble(procKach1);

                    if (name == MEGA[0][i])
                    {

                        bool[] Amega = new bool[6];
                        int d = 0;
                        Amega[d++] = Convert.ToDouble(MEGA[c][2]) > Convert.ToDouble(podMEGA1);
                        Amega[d++] = Convert.ToDouble(MEGA[c][12]) > Convert.ToDouble(sredPop1);
                        Amega[d++] = Convert.ToDouble(MEGA[c][3]) > Convert.ToDouble(m2act1);
                        Amega[d++] = Convert.ToDouble(MEGA[c][111]) > Convert.ToDouble(procKach1);

                        if ()
                            NEW[c, 2] = "green";
                        //...
                        find = true;
                        break;
                    }
                }


                if (!find)
                {
                    if (Convert.ToInt32(MTS[4][c]) > Convert.ToInt32(blueZone)) // Allincom
                    {
                        //...
                        NEW[c, 4] += "продает только МТС /n";
                        if (Convert.ToDouble(MTS[c][2]) > Convert.ToInt32(podMTS1))
                        {
                            NEW[c, 2] = "darkgreen";
                        }
                        else if ((Convert.ToDouble(MTS[c][2]) < Convert.ToInt32(podMTS2)))
                        {
                            NEW[c, 2] = "red";
                        }
                        else
                        {
                            NEW[c, 2] = "yellow";
                        }
                    }
                    else
                    {
                        NEW[c, 2] = "blue";
                    }
                }
            }


        }
    }
}
