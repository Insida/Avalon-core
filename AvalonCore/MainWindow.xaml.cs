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
using System.Data.SqlClient;

namespace AvalonCore
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        const string conn = @"Server=(localdb)\mssqllocaldb;database=AvalonVR;";
        public MainWindow()
        {
            InitializeComponent();
            FillMainGrid();
        }

        public class Order
        {
            public string client { get; set; }
            public string zone { get; set; }
            public string playtime { get; set; }
            public string orderdesc { get; set; }
        }

        private void DataGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        public AvalonVRDS AVRDS = new AvalonVRDS(); // rename DataSet to use in future


        private void Button_Click(object sender, RoutedEventArgs e)
        {
            this.WindowState = WindowState.Minimized;
        }
        private void Games_Click(object sender, RoutedEventArgs e)
        {
            MainWindowGrid.Opacity = 0; MainWindowGrid.Visibility = Visibility.Hidden;
            GamesGrid.Opacity = 100; GamesGrid.Visibility = Visibility.Visible;
        }
        private void Main_Click(object sender, RoutedEventArgs e)
        {
            MainWindowGrid.Opacity = 100; MainWindowGrid.Visibility = Visibility.Visible;
            GamesGrid.Opacity = 0; GamesGrid.Visibility = Visibility.Hidden;
            
        }
        public void FillMainGrid()  // Заполнение Грида значениями датасета
        {
            DataGridTextColumn col1 = new DataGridTextColumn();
            DataGridTextColumn col2 = new DataGridTextColumn();
            DataGridTextColumn col3 = new DataGridTextColumn();
            DataGridTextColumn col4 = new DataGridTextColumn();
            col1.Header = "ФИО"; col1.Binding = new Binding("client"); col1.Width = 260;
            DGV1.Columns.Add(col1);
            col2.Header = "Зона"; col2.Binding = new Binding("zone"); col2.Width = 260;
            DGV1.Columns.Add(col2);
            col3.Header = "Время"; col3.Binding = new Binding("playtime"); col3.Width = 260;
            DGV1.Columns.Add(col3); 
            col4.Header = "Описание"; col4.Binding = new Binding("orderdesc"); col4.Width = 269;
            DGV1.Columns.Add(col4);


            //DGV1.RowBackground = Brushes.Gray;
            //DGV1.AlternatingRowBackground = Brushes.Gray;
            //DGV1.ColumnHeaderHeight = 0;
            //DGV1.RowHeaderWidth = 0;

            try
            {
                SqlConnection con = new SqlConnection(conn);
                SqlConnection con1 = new SqlConnection(conn);
                con.Open();
                string get = "Select zonename from zones"; // Заполнение CB1
                SqlCommand cmd = new SqlCommand(get, con);
                SqlDataReader reader = cmd.ExecuteReader();
                if (reader.HasRows)
                {
                    while (reader.Read())
                    {
                        CB1.Items.Add(reader.GetValue(0).ToString());
                    }
                }
                con.Close(); con.Open(); // Переоткрытие соедениния

                get = "Select gamename from games"; // Заполнение CB2
                cmd = new SqlCommand(get, con);
                reader = cmd.ExecuteReader();
                if (reader.HasRows)
                {
                    while (reader.Read())
                    {
                        CB2.Items.Add(reader.GetValue(0).ToString());
                    }
                }


                con.Close(); con.Open();

                get = "Select * from orders"; // Заполнение CB2
                string getcbyid, getzbyid, time, desc, fio, zone;
                cmd = new SqlCommand(get, con);
                reader = cmd.ExecuteReader();
                if (reader.HasRows)
                {
                    while (reader.Read())
                    {
                        getcbyid = "Select fio from clients where clientid=" + reader.GetValue(1).ToString();
                        getzbyid = "Select zonename from zones where zoneid=" + reader.GetValue(2).ToString();
                        time = reader.GetValue(3).ToString();
                        desc = reader.GetValue(4).ToString();
                        con1.Close(); con1.Open();
                        cmd = new SqlCommand(getcbyid, con1);
                        object getfio = cmd.ExecuteScalar();
                        fio = getfio.ToString();
                        con1.Close(); con1.Open();
                        cmd = new SqlCommand(getzbyid, con1);
                        object getzone = cmd.ExecuteScalar();
                        zone = getzone.ToString();
                        DGV1.Items.Add(new Order() {client = fio,zone = zone,playtime=time,orderdesc=desc });
                    }
                }
            }
            catch (SqlException ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public void FillGamesGrid()
        {
            
        }


        private void TryToFillClient(object sender, TextChangedEventArgs e)
        {
            string fio = TB1.Text;
            SqlConnection con = new SqlConnection(conn);
            con.Open();
            string get = "Select FIO,num from clients"; // Заполнение CB1
            SqlCommand cmd = new SqlCommand(get, con);
            SqlDataReader reader = cmd.ExecuteReader();
            if (reader.HasRows)
            {
                while (reader.Read())
                {
                    if (reader.GetValue(0).ToString().StartsWith(fio) && TB1.Text.Length>3) // Автозаполнение при нажатом Enter
                    {
                        labelfio.Content = reader.GetValue(0).ToString();
                        if(Keyboard.IsKeyDown(Key.Enter))
                        {
                            TB1.Text = labelfio.Content.ToString();
                            labelfio.Content = "";
                        }
                        if (reader.GetValue(0).ToString().Equals(fio))
                        {
                            TBNum.Text = reader.GetValue(1).ToString();
                        }
                    }
                    else
                    {
                        labelfio.Content = "";
                    }
                }
            }
            con.Close();
        }
    }
}
