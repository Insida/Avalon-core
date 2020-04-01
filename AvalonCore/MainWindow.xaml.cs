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
using System.IO;


using System.Data.SqlClient;
using Word = Microsoft.Office.Interop.Word;
using System.Reflection;


namespace AvalonCore
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        Encoding enc = Encoding.GetEncoding(1251);
        const string conn = @"Server=(localdb)\mssqllocaldb;database=AvalonVR;";
        public MainWindow()
        {
            InitializeComponent();
            FillMainGrid();
            FillGamesGrid();
            FillZonesGrid();
            UsersFill();
            FillPartnersGrid();
            FillSpecialOrdersGrid();
        }

        public class Order
        {
            public string client { get; set; }
            public string zone { get; set; }
            public string playtime { get; set; }
            public string orderdesc { get; set; }
            public string date { get; set; }
            public string price { get; set; }
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
        public class User
        {
            public string username { get; set; }
            public string usertypeid { get; set; }
        }
        public class Partner
        {
            public string clientid { get; set; }
            public string partnername { get; set; }
            public string adress { get; set; }
            public string bank { get; set; }
            public string banknum { get; set; }
            public string unp { get; set; }
        }
        public class specialorder
        {
            public string partnersid { get; set; }
            public string specialorderdesc { get; set; }
            public string specialordertime { get; set; }
            public string specialorderdate { get; set; }
            public string price { get; set; }
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
            GamesGrid.Visibility = Visibility.Visible;
            SpecialOrdersGrid.Visibility = Visibility.Hidden;
            PartnersGrid.Visibility = Visibility.Hidden;
            MainWindowGrid.Visibility = Visibility.Hidden;
            ZonesGrid.Visibility = Visibility.Hidden;
            UsersGrid.Visibility = Visibility.Hidden;
        }
        private void Main_Click(object sender, RoutedEventArgs e)
        {
            MainWindowGrid.Visibility = Visibility.Visible;
            SpecialOrdersGrid.Visibility = Visibility.Hidden;
            PartnersGrid.Visibility = Visibility.Hidden;
            GamesGrid.Visibility = Visibility.Hidden;
            ZonesGrid.Visibility = Visibility.Hidden;
            UsersGrid.Visibility = Visibility.Hidden;
        }
        private void Zone_Click(object sender, RoutedEventArgs e)
        {
            ZonesGrid.Visibility = Visibility.Visible;
            SpecialOrdersGrid.Visibility = Visibility.Hidden;
            PartnersGrid.Visibility = Visibility.Hidden;
            MainWindowGrid.Visibility = Visibility.Hidden;
            GamesGrid.Visibility = Visibility.Hidden;
            UsersGrid.Visibility = Visibility.Hidden;
        }
        private void UsersClick(object sender, RoutedEventArgs e)
        {
            UsersGrid.Visibility = Visibility.Visible;
            SpecialOrdersGrid.Visibility = Visibility.Hidden;
            PartnersGrid.Visibility = Visibility.Hidden;
            MainWindowGrid.Visibility = Visibility.Hidden;
            GamesGrid.Visibility = Visibility.Hidden;
            ZonesGrid.Visibility = Visibility.Hidden;
        }
        private void SpecialClick(object sender, RoutedEventArgs e)
        {
            SpecialOrdersGrid.Visibility = Visibility.Visible;
            PartnersGrid.Visibility = Visibility.Hidden;
            UsersGrid.Visibility = Visibility.Hidden;
            MainWindowGrid.Visibility = Visibility.Hidden;
            GamesGrid.Visibility = Visibility.Hidden;
            ZonesGrid.Visibility = Visibility.Hidden;
        }
        private void PartnersClick(object sender, RoutedEventArgs e)
        {
            PartnersGrid.Visibility = Visibility.Visible;
            SpecialOrdersGrid.Visibility = Visibility.Hidden;
            UsersGrid.Visibility = Visibility.Hidden;
            MainWindowGrid.Visibility = Visibility.Hidden;
            GamesGrid.Visibility = Visibility.Hidden;
            ZonesGrid.Visibility = Visibility.Hidden;
        }



        public void FillMainGrid()  // Заполнение Грида значениями датасета
        {
            DataGridTextColumn col1 = new DataGridTextColumn();
            DataGridTextColumn col2 = new DataGridTextColumn();
            DataGridTextColumn col3 = new DataGridTextColumn();
            DataGridTextColumn col4 = new DataGridTextColumn();
            DataGridTextColumn col5 = new DataGridTextColumn();
            DataGridTextColumn col6 = new DataGridTextColumn();


            col1.Header = "ФИО"; col1.Binding = new Binding("client"); col1.Width = 174;
            DGV1.Columns.Add(col1);
            col2.Header = "Зона"; col2.Binding = new Binding("zone"); col2.Width = 174;
            DGV1.Columns.Add(col2);
            col3.Header = "Время"; col3.Binding = new Binding("playtime"); col3.Width = 174;
            DGV1.Columns.Add(col3);
            col4.Header = "Описание"; col4.Binding = new Binding("orderdesc"); col4.Width = 174;
            DGV1.Columns.Add(col4);
            col5.Header = "Дата"; col5.Binding = new Binding("date"); col5.Width = 174;
            DGV1.Columns.Add(col5);
            col6.Header = "Цена"; col6.Binding = new Binding("price"); col6.Width = 174;
            DGV1.Columns.Add(col6);



            DGV1.MaxColumnWidth = 174; DGV1.MinColumnWidth = 174;


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

                get = "Select * from orders where date = '" + NowSQL()+"'"; // Заполнение CB2
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
                        DGV1.Items.Add(new Order() { client = fio, zone = zone, playtime = time, orderdesc = desc, date = reader.GetValue(5).ToString().Remove(10), price = reader.GetValue(6).ToString() });
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
                if (reader.HasRows)
                {
                    while (reader.Read())
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
                string gettypes = "Select zonetypename from zonetypes";
                SqlCommand cmd = new SqlCommand(gettypes, con);
                SqlDataReader reader = cmd.ExecuteReader();
                if (reader.HasRows)
                {
                    while (reader.Read())
                    {
                        ZTCB.Items.Add(reader.GetValue(0).ToString());
                    }
                }


                con.Close(); con.Open();



                string getgames = "Select zonename,zonetypeid,tenminprice,thirtyminprice,sixtyminprice from zones";
                cmd = new SqlCommand(getgames, con);
                reader = cmd.ExecuteReader();
                if (reader.HasRows)
                {
                    while (reader.Read())
                    {
                        string getbyzoneid = "Select zonetypename from zonetypes where zonetypeid=" + reader.GetValue(1).ToString();
                        con1.Close(); con1.Open();
                        cmd = new SqlCommand(getbyzoneid, con1);
                        object getzonetype = cmd.ExecuteScalar();
                        ZonesDGV.Items.Add(new Zone { zonename = reader.GetValue(0).ToString(), zonetypeid = getzonetype.ToString(), tenminprice = reader.GetValue(2).ToString(),
                            thirtyminprice = reader.GetValue(3).ToString(), sixtyminprice = reader.GetValue(4).ToString() });
                    }
                }

            }
            catch (SqlException ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void UsersFill()
        {
            DataGridTextColumn col1 = new DataGridTextColumn();
            DataGridTextColumn col2 = new DataGridTextColumn();
            col1.Header = "Пользователь"; col1.Binding = new Binding("username"); col1.Width = 522;
            UsersDGV.Columns.Add(col1);
            col2.Header = "Тип"; col2.Binding = new Binding("usertypeid"); col2.Width = 522;
            UsersDGV.Columns.Add(col2);
            try
            {
                SqlConnection con = new SqlConnection(conn);
                SqlConnection con1 = new SqlConnection(conn);
                con.Open();
                string get = "Select usertypename from usertypes"; // Заполнение CB
                SqlCommand cmd = new SqlCommand(get, con);
                SqlDataReader reader = cmd.ExecuteReader();
                if (reader.HasRows)
                {
                    while (reader.Read())
                    {
                        UTCB.Items.Add(reader.GetValue(0).ToString());
                    }
                }
                con.Close(); con.Open();
                get = "select * from users";
                cmd = new SqlCommand(get, con);
                reader = cmd.ExecuteReader();
                if (reader.HasRows)
                {
                    while (reader.Read())
                    {
                        con1.Close(); con1.Open();
                        string usertype = "select usertypename from usertypes where usertypeid=" + reader.GetValue(3).ToString();
                        cmd = new SqlCommand(usertype, con1);
                        object getusertype = cmd.ExecuteScalar();
                        UsersDGV.Items.Add(new User { username = reader.GetValue(1).ToString(), usertypeid = getusertype.ToString() });
                    }
                }
            }
            catch (SqlException ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void FillPartnersGrid()
        {
            DataGridTextColumn col1 = new DataGridTextColumn();
            DataGridTextColumn col2 = new DataGridTextColumn();
            DataGridTextColumn col3 = new DataGridTextColumn();
            DataGridTextColumn col4 = new DataGridTextColumn();
            DataGridTextColumn col5 = new DataGridTextColumn();
            DataGridTextColumn col6 = new DataGridTextColumn();
            col1.Header = "Наименование"; col1.Binding = new Binding("partnername"); col1.Width = 174;
            PartnersDGV.Columns.Add(col1);
            col2.Header = "Ответ.Лицо"; col2.Binding = new Binding("clientid"); col2.Width = 174;
            PartnersDGV.Columns.Add(col2);
            col3.Header = "Адрес"; col3.Binding = new Binding("adress"); col3.Width = 174;
            PartnersDGV.Columns.Add(col3);
            col4.Header = "Банк"; col4.Binding = new Binding("bank"); col4.Width = 174;
            PartnersDGV.Columns.Add(col4);
            col5.Header = "Номер счета"; col5.Binding = new Binding("banknum"); col5.Width = 174;
            PartnersDGV.Columns.Add(col5);
            col6.Header = "УНП"; col6.Binding = new Binding("unp"); col6.Width = 174;
            PartnersDGV.Columns.Add(col6);
            BankCB.Items.Add("Идея Банк"); BankCB.Items.Add("БелВЭБ"); BankCB.Items.Add("Решение"); BankCB.Items.Add("Дабрабыт"); BankCB.Items.Add("Абсолютбанк"); BankCB.Items.Add("Альфа-Банк"); BankCB.Items.Add("БПС-Сбербанк"); BankCB.Items.Add("БСБ");
            BankCB.Items.Add("БТА"); BankCB.Items.Add("ВТБ"); BankCB.Items.Add("БелГазпромБанк"); BankCB.Items.Add("БелАгроПромБанк"); BankCB.Items.Add("БеларусБанк"); BankCB.Items.Add("ББМБ"); BankCB.Items.Add("БНБ"); BankCB.Items.Add("РРБ-Банк");
            BankCB.Items.Add("МТБанк"); BankCB.Items.Add("Статусбанк"); BankCB.Items.Add("ФрансаБанк"); BankCB.Items.Add("ТК"); BankCB.Items.Add("Хоум Кредит"); BankCB.Items.Add("ТехноБанк"); BankCB.Items.Add("ЕвроБанк"); BankCB.Items.Add("Дельта");
            BankCB.Items.Add("ИнтерПэйБанк"); BankCB.Items.Add("Паритетбанк"); BankCB.Items.Add("НБРБ"); BankCB.Items.Add("ПриорБанк"); BankCB.Items.Add("Цептер");
            try
            {
                SqlConnection con = new SqlConnection(conn);
                SqlConnection con1 = new SqlConnection(conn);
                con.Open();
                string get = "select * from partners";
                SqlCommand cmd = new SqlCommand(get, con);
                SqlDataReader reader = cmd.ExecuteReader();
                if (reader.HasRows)
                {
                    while (reader.Read())
                    {
                        con1.Close(); con1.Open();
                        string usertype = "select fio from clients where clientid=" + reader.GetValue(1).ToString();
                        cmd = new SqlCommand(usertype, con1);
                        object getusertype = cmd.ExecuteScalar();
                        PartnersDGV.Items.Add(new Partner {clientid=getusertype.ToString(), partnername=reader.GetValue(2).ToString(), adress = reader.GetValue(3).ToString(), bank = reader.GetValue(4).ToString(), banknum = reader.GetValue(5).ToString(),
                            unp = reader.GetValue(6).ToString() });
                    }
                }
            }
            catch (SqlException ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        public void FillSpecialOrdersGrid()
        {
            DataGridTextColumn col1 = new DataGridTextColumn();
            DataGridTextColumn col2 = new DataGridTextColumn();
            DataGridTextColumn col3 = new DataGridTextColumn();
            DataGridTextColumn col4 = new DataGridTextColumn();
            DataGridTextColumn col5 = new DataGridTextColumn();
            col1.Header = "Партнер"; col1.Binding = new Binding("partnersid"); col1.Width = 208;
            SpecialOrdersDGV.Columns.Add(col1);
            col2.Header = "Описание"; col2.Binding = new Binding("specialorderdesc"); col2.Width = 208;
            SpecialOrdersDGV.Columns.Add(col2);
            col3.Header = "Время"; col3.Binding = new Binding("specialordertime"); col3.Width = 208;
            SpecialOrdersDGV.Columns.Add(col3);
            col4.Header = "Дата"; col4.Binding = new Binding("specialorderdate"); col4.Width = 208;
            SpecialOrdersDGV.Columns.Add(col4);
            col5.Header = "Цена"; col5.Binding = new Binding("price"); col5.Width = 208;
            SpecialOrdersDGV.Columns.Add(col5);
            try
            {
                SqlConnection con = new SqlConnection(conn);
                SqlConnection con1 = new SqlConnection(conn);
                con.Open();
                string get = "select * from specialorders";
                SqlCommand cmd = new SqlCommand(get, con);
                SqlDataReader reader = cmd.ExecuteReader();
                if (reader.HasRows)
                {
                    while (reader.Read())
                    {
                        con1.Close(); con1.Open();
                        string usertype = "select partnername from partners where partnerid=" + reader.GetValue(1).ToString();
                        cmd = new SqlCommand(usertype, con1);
                        object getusertype = cmd.ExecuteScalar();
                        SpecialOrdersDGV.Items.Add(new specialorder
                        {
                            partnersid = getusertype.ToString(),
                            specialorderdesc = reader.GetValue(2).ToString(),
                            specialordertime = reader.GetValue(3).ToString(),
                            specialorderdate = reader.GetValue(4).ToString().Remove(10),
                            price = reader.GetValue(5).ToString()
                        });
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
                    if (reader.GetValue(0).ToString().StartsWith(fio) && TB1.Text.Length > 3) // Автозаполнение при нажатом Enter
                    {
                        labelfio.Content = reader.GetValue(0).ToString();
                        if (Keyboard.IsKeyDown(Key.Enter))
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

        private void AddOrderButton(object sender, RoutedEventArgs e)
        {
            if (TB1.Text.Length != 0 && TBNum.Text.Length != 0 && CB1.Text.Length != 0 && CBTime.Text.Length != 0 && CB2.Text.Length != 0 && TBOrderDesc.Text.Length != 0)
            {
                try
                {
                    SqlConnection con = new SqlConnection(conn);
                    con.Open();
                    string strsql = "if 0=(select count(num) from clients where num = '" + TBNum.Text + "') if 0=(select count(fio) from clients where fio = '" + TB1.Text + "') INSERT INTO [clients] VALUES(" + "'" + TB1.Text + "','" + TBNum.Text + "')";
                    SqlCommand cmd = new SqlCommand(strsql, con);
                    if (cmd.ExecuteNonQuery() == 1)
                        MessageBox.Show("Запись успешно добавлена.");
                    con.Close();
                }
                catch (SqlException ex)
                {
                    MessageBox.Show(ex.Message);
                }
                try
                {
                    SqlConnection con = new SqlConnection(conn);
                    con.Open();
                    //string strsql = "select clientid from clients where fio = '" + TB1.Text + "'"+"select zoneid from zones where zonename = '" + CB1.Text + "'";
                    string strsql = "select clientid from clients where fio = '" + TB1.Text + "'";
                    string getzid = "select zoneid,tenminprice,thirtyminprice,sixtyminprice from zones where zonename = '" + CB1.Text + "'";
                    SqlCommand cmd = new SqlCommand(strsql, con);
                    SqlCommand cmd1 = new SqlCommand(getzid, con);
                    object cid = cmd.ExecuteScalar();
                    SqlDataReader zid = cmd1.ExecuteReader();
                    MessageBox.Show(cid.ToString());
                    if (zid.HasRows)
                        while (zid.Read())
                        {
                            string price = "0";
                            if (CBTime.Text.ToString().Remove(2) == "10")
                                price = zid.GetValue(1).ToString();
                            else
                                if (CBTime.Text.ToString().Remove(2) == "30")
                                price = zid.GetValue(2).ToString();
                            else
                                price = zid.GetValue(3).ToString();
                            strsql = "INSERT INTO orders VALUES('" + cid.ToString() + "','" + zid.GetValue(0).ToString() + "','" + CBTime.Text.ToString().Remove(2) + "','" + TBOrderDesc.Text + " " + CB2.Text.ToString() + "','" + DateTime.Now.Date.ToString().Replace('.', '-') + "'" + price + "')";
                        }
                    con.Close(); con.Open();
                    cmd = new SqlCommand(strsql, con);
                    if (cmd.ExecuteNonQuery() == 1)
                    {
                        MessageBox.Show("Запись успешно добавлена.");
                        //// Добавить Чек
                        //var app = new Word.Application();
                        //app.Visible = true;
                        //var doc = app.Documents.Add();
                        //var r = doc.Range();
                        //r.Text = "Avalon-VR/n"+DateTime.Now+"/n Сумма = "+CB2.Text.ToString()+"/n ЦБ_РБ";
                        //doc.SaveAs (@"D:\Check "+DateTime.Now);
                    }
                    con.Close();
                }
                catch (SqlException ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            else MessageBox.Show("Неверно введены данные");
        }

        private void AddGamesButton(object sender, RoutedEventArgs e)
        {
            try
            {
                SqlConnection con = new SqlConnection(conn);
                con.Open();
                string strsql = "if 0=(select count(gamename) from games where gamename = '" + GNTB.Text + "') INSERT INTO [games] VALUES(" + "'" + GNTB.Text + "','" + GDTB.Text + "')";
                SqlCommand cmd = new SqlCommand(strsql, con);
                if (cmd.ExecuteNonQuery() == 1)
                    MessageBox.Show("Запись успешно добавлена.");
                con.Close();
            }
            catch (SqlException ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void ReloadMainClick(object sender, RoutedEventArgs e)
        {
            DGV1.Columns.Clear();
            DGV1.Items.Clear();
            CB1.Items.Clear();
            CB2.Items.Clear();
            FillMainGrid();
        }

        private void ReloadGamesButton(object sender, RoutedEventArgs e)
        {
            GamesDGV.Columns.Clear();
            GamesDGV.Items.Clear();
            FillGamesGrid();
        }

        private string DateToSQLFormate(string date)
        {
            string day = date.Remove(2);
            string month = date[3].ToString() + date[4].ToString();
            string year = date[6].ToString() + date[7].ToString() + date[8].ToString() + date[9].ToString();
            date = year+"-"+month+"-"+day;
            return date;
        }

        private string NowSQL()
        {
            string Nowadate = DateTime.Now.Date.Year + "-" + DateTime.Now.Date.Month + "-" + DateTime.Now.Date.Day;
            return Nowadate;
        }

        private void ReportButtonClick(object sender, RoutedEventArgs e)
        {
            // Отчет
            var app = new Word.Application();
            app.Visible = true;
            var doc = app.Documents.Add();
            var r = doc.Range();
            r.Text = "Avalon-VR" + " Дата начала = " + DPStart.Text.ToString() + " Дата конца= "+DPEnd.Text.ToString();
            string get = "Select * from orders where date > '" + DateToSQLFormate(DPStart.Text.ToString()).ToString() + "' and date < '"+ DateToSQLFormate(DPEnd.Text.ToString()).ToString()+"'";
            try
            {
                SqlConnection con = new SqlConnection(conn);
                SqlConnection con1 = new SqlConnection(conn);
                con.Open();
                get = "Select * from orders where date > '" + DateToSQLFormate(DPStart.Text.ToString()).ToString() + "' and date < '" + DateToSQLFormate(DPEnd.Text.ToString()).ToString() + "'"; ;
                string getcbyid, getzbyid, time, desc, fio, zone;
                SqlCommand cmd = new SqlCommand(get, con);
                SqlDataReader reader = cmd.ExecuteReader();
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
                        r.Text += " ФИО " + fio + " Зона " + zone + " Время " + time + " Описание " + desc + " Дата " + reader.GetValue(5).ToString().Remove(10) + " Цена " + reader.GetValue(6).ToString();
                    }
                }
            }
            catch (SqlException ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }
}
