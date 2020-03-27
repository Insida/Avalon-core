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
            FillGamesGrid();
            FillZonesGrid();
        }

        public class Order
        {
            public string client { get; set; }
            public string zone { get; set; }
            public string playtime { get; set; }
            public string orderdesc { get; set; }
        }
        public class Game
        {
            public string gamename { get; set; }
            public string gamedesc { get; set; }
        }
        public class Zone
        {
            public string zonename { get; set; }
            public string zonetypeid { get; set; }
            public string tenminprice { get; set; }
            public string thirtyminprice { get; set; }
            public string sixtyminprice { get; set; }
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            this.WindowState = WindowState.Minimized;
        }

        public AvalonVRDS AVRDS = new AvalonVRDS(); // rename DataSet to use in future


        private void Games_Click(object sender, RoutedEventArgs e)
        {
            MainWindowGrid.Visibility = Visibility.Hidden;
            GamesGrid.Visibility = Visibility.Visible;
            ZonesGrid.Visibility = Visibility.Hidden;
        }
        private void Main_Click(object sender, RoutedEventArgs e)
        {
            MainWindowGrid.Visibility = Visibility.Visible;
            GamesGrid.Visibility = Visibility.Hidden;
            ZonesGrid.Visibility = Visibility.Hidden;
        }
        private void Zone_Click(object sender, RoutedEventArgs e)
        {
            MainWindowGrid.Visibility = Visibility.Hidden;
            GamesGrid.Visibility = Visibility.Hidden;
            ZonesGrid.Visibility = Visibility.Visible;
        }


        public void FillMainGrid()  // Заполнение Грида значениями датасета
        {
            DataGridTextColumn col1 = new DataGridTextColumn();
            DataGridTextColumn col2 = new DataGridTextColumn();
            DataGridTextColumn col3 = new DataGridTextColumn();
            DataGridTextColumn col4 = new DataGridTextColumn();
            col1.Header = "ФИО"; col1.Binding = new Binding("client"); col1.Width = 261;
            DGV1.Columns.Add(col1);
            col2.Header = "Зона"; col2.Binding = new Binding("zone"); col2.Width = 261;
            DGV1.Columns.Add(col2);
            col3.Header = "Время"; col3.Binding = new Binding("playtime"); col3.Width = 261;
            DGV1.Columns.Add(col3); 
            col4.Header = "Описание"; col4.Binding = new Binding("orderdesc"); col4.Width = 261;
            DGV1.Columns.Add(col4);
            DGV1.MaxColumnWidth = 261; DGV1.MinColumnWidth = 261;


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
            DataGridTextColumn col1 = new DataGridTextColumn();
            DataGridTextColumn col2 = new DataGridTextColumn();
            col1.Header = "Название"; col1.Binding = new Binding("gamename"); col1.Width = 522;
            GamesDGV.Columns.Add(col1);
            col2.Header = "Описание"; col2.Binding = new Binding("gamedesc"); col2.Width = 522;
            GamesDGV.Columns.Add(col2);
            try
            {
                SqlConnection con = new SqlConnection(conn);
                con.Open();
                string getgames = "Select gamename,gamedesc from games";
                SqlCommand cmd = new SqlCommand(getgames, con);
                SqlDataReader reader = cmd.ExecuteReader();
                if(reader.HasRows)
                {
                    while(reader.Read())
                    {
                        GamesDGV.Items.Add(new Game { gamename = reader.GetValue(0).ToString(), gamedesc = reader.GetValue(1).ToString() });
                    }     
                }
                
            }
            catch (SqlException ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        public void FillZonesGrid()
        {
            DataGridTextColumn col1 = new DataGridTextColumn();
            DataGridTextColumn col2 = new DataGridTextColumn();
            DataGridTextColumn col3 = new DataGridTextColumn();
            DataGridTextColumn col4 = new DataGridTextColumn();
            DataGridTextColumn col5 = new DataGridTextColumn();
            col1.Header = "Название"; col1.Binding = new Binding("zonename"); col1.Width = 210;
            ZonesDGV.Columns.Add(col1);
            col2.Header = "Тип"; col2.Binding = new Binding("zonetypeid"); col2.Width = 210;
            ZonesDGV.Columns.Add(col2);
            col3.Header = "10 мин"; col3.Binding = new Binding("tenminprice"); col3.Width = 210;
            ZonesDGV.Columns.Add(col3);
            col4.Header = "30 мин"; col4.Binding = new Binding("thirtyminprice"); col4.Width = 210;
            ZonesDGV.Columns.Add(col4);
            col5.Header = "60 мин"; col5.Binding = new Binding("sixtyminprice"); col5.Width = 210;
            ZonesDGV.Columns.Add(col5);
            DGV1.MaxColumnWidth = 210; DGV1.MinColumnWidth = 210;



            try
            {
                SqlConnection con = new SqlConnection(conn);
                SqlConnection con1 = new SqlConnection(conn);
                con.Open();
                string getgames = "Select zonename,zonetypeid,tenminprice,thirtyminprice,sixtyminprice from zones";
                SqlCommand cmd = new SqlCommand(getgames, con);
                SqlDataReader reader = cmd.ExecuteReader();
                if (reader.HasRows)
                {
                    while (reader.Read())
                    {
                        string getbyzoneid = "Select zonetypename from zonetypes where zonetypeid=" + reader.GetValue(1).ToString();
                        con1.Open();
                        cmd = new SqlCommand(getbyzoneid, con1);
                        object getzonetype = cmd.ExecuteScalar();
                        ZonesDGV.Items.Add(new Zone { zonename = reader.GetValue(0).ToString(), zonetypeid =getzonetype.ToString(), tenminprice = reader.GetValue(2).ToString(),
                            thirtyminprice = reader.GetValue(3).ToString(), sixtyminprice = reader.GetValue(4).ToString() });
                    }
                }

            }
            catch (SqlException ex)
            {
                MessageBox.Show(ex.Message);
            }
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
