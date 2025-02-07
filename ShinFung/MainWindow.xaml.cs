﻿using System;
using System.Windows;
using System.IO;
using Microsoft.Data.Sqlite;
using System.Diagnostics;
using System.Collections.Generic;
using System.ComponentModel;
using System.Windows.Controls;
using System.Windows.Data;
using System.Globalization;
using System.Collections.ObjectModel;
using System.Windows.Input;
using ShinFung.Properties;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Path = System.IO.Path;
using static System.Net.Mime.MediaTypeNames;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Documents;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Wordprocessing;
using CheckBox = System.Windows.Controls.CheckBox;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
using Text = DocumentFormat.OpenXml.Wordprocessing.Text;
using System.Security.Cryptography;
using DocumentFormat.OpenXml.ExtendedProperties;
using System.Threading;

namespace ShinFung
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window, INotifyPropertyChanged
    {
        string DocumentPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        string DbFileName = "db2.db";
        string DataFolder = "ShinFung";
        ObservableCollection<ListItem> Items = new ObservableCollection<ListItem>();
        public ObservableCollection<ListItem> Data {
            get { return Items; }
            set 
            { 
                Items = value;
                this.OnPropertyChanged("Data");
            }
        }
        public event PropertyChangedEventHandler? PropertyChanged;
        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChangedEventHandler handler = PropertyChanged;
            if (handler != null)
            {
                handler(this, new PropertyChangedEventArgs(propertyName));
            }
        }
        public MainWindow()
        {
            InitializeComponent();
            CreateDataBase();
            BaseBindings();
            ReadAllFromDbThread();
        }
        private void BaseBindings()
        {
            CheckBox_S_C.SetBinding(CheckBox.IsCheckedProperty,new Binding("Visibility") { Source = this.ListBox_S_C, Converter = new BoolToVisibilityConverter(), Mode = BindingMode.OneWayToSource});
            MainList.SetBinding(ListView.ItemsSourceProperty, new Binding("Data") { Source = this });
        }
        private void CreateDataBase()
        {
            string DbFilePath = Path.Combine(DocumentPath, DataFolder, DbFileName);
            string _connectionString = "Data Source="+DbFilePath;
            if (File.Exists(DbFilePath)) return;//file exist cancel creation
            Directory.CreateDirectory(Path.Combine(DocumentPath, DataFolder));
            using (var connection = new SqliteConnection(_connectionString))
            {
                connection.Open();
                var command = connection.CreateCommand();
                //YYYY-MM-DD HH:MM:SS.SSS
                string tableString = SqlStruct.BaseStruct;
                command.CommandText = tableString;
                command.ExecuteNonQuery();
            }


        }

        private void AddData(object sender, RoutedEventArgs e)
        {
            Insert();
            ReadAllFromDb();
        }
        //ListData
        private void ListData(object sender, RoutedEventArgs e)
        {
            ReadAllFromDb();
        }

        private void ClearItems()
        {
            this.Dispatcher.Invoke(() =>
            {
                Items.Clear();
            });
        }

        private void ReadAllFromDb()
        {
            ClearItems();
            string DbFilePath = Path.Combine(DocumentPath, DataFolder, DbFileName);
            string _connectionString = "Data Source=" + DbFilePath;
            using (var connection = new SqliteConnection(_connectionString))
            {
                connection.Open();
                var command = connection.CreateCommand();
                command.CommandText = @"  select * from userdata WHERE valid > 0 ORDER BY id DESC";

                using (var reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        object[] datas = new object[reader.FieldCount];
                        reader.GetValues(datas);
                        AddList(new ListItem(datas));
                    }
                }
            }
        }

        private void CalculateTotalSum() 
        {
            int total = 0;
            foreach( ListItem item in Items)
            {
                total += int.Parse(item.Sum);
            }
        }

        private void ReadAllFromDbThread()
        {
            // Start a new background thread for the long-running task
            Thread backgroundThread = new Thread(ReadAllFromDb);
            backgroundThread.Start();
        }

        private void EditSelected()
        {
            string DbFilePath = Path.Combine(DocumentPath, DataFolder, DbFileName);
            string _connectionString = "Data Source=" + DbFilePath;
            ListItem select = MainList.SelectedItem as ListItem;

            if (select == null)
                return;

            using (var connection = new SqliteConnection(_connectionString))
            {
                connection.Open();
                var command = connection.CreateCommand();
                try
                {
                    int id;
                    if (!int.TryParse(select.Id, out id))
                    {
                        throw new InvalidOperationException("Invalid ID: ID must be a valid integer.");
                    }

                    command.CommandText = @"
                                            UPDATE userdata 
                                            SET 
                                                name = $name,
                                                zodiac = $zodiac,
                                                age = $age,
                                                S_A = $S_A,
                                                S_B = $S_B,
                                                S_C = $S_C,
                                                S_C_type = $S_C_type,
                                                S_D = $S_D,
                                                S_E = $S_E,
                                                address = $address,
                                                D_A = $D_A,
                                                D_B = $D_B,
                                                phonenumber = $phonenumber,
                                                datetime = $datetime,
                                                sum = $sum,
                                                target_name = $target_name,
                                                employee_name = $employee_name,
                                                valid = $valid
                                            WHERE id = $id;
                                        ";

                    command.Parameters.AddWithValue("$id", id);
                    command.Parameters.AddWithValue("$name", TextBox_Name.Text.Length > 0 ? TextBox_Name.Text : string.Empty);
                    command.Parameters.AddWithValue("$zodiac", TextBox_Zodiac.Text.Length > 0 ? TextBox_Zodiac.Text : string.Empty);
                    command.Parameters.AddWithValue("$age", TextBox_Age.Text.Length > 0 ? int.Parse(TextBox_Age.Text) : 0);
                    command.Parameters.AddWithValue("$S_A", (CheckBox_S_A.IsChecked ?? false) ? int.Parse(TextBox_S_A.Text) : 0);
                    command.Parameters.AddWithValue("$S_B", (CheckBox_S_B.IsChecked ?? false) ? int.Parse(TextBox_S_B.Text) : 0);
                    command.Parameters.AddWithValue("$S_C", (CheckBox_S_C.IsChecked ?? false) ? int.Parse(TextBox_S_C.Text) : 0);
                    command.Parameters.AddWithValue("$S_C_type", (CheckBox_S_C.IsChecked ?? false) ? ListBox_S_C.SelectedIndex : -1);
                    command.Parameters.AddWithValue("$S_D", (CheckBox_S_D.IsChecked ?? false) ? int.Parse(TextBox_S_D.Text) : 0);
                    command.Parameters.AddWithValue("$S_E", (CheckBox_S_E.IsChecked ?? false) ? int.Parse(TextBox_S_E.Text) : 0);
                    command.Parameters.AddWithValue("$address", TextBox_Address.Text.Length > 0 ? TextBox_Address.Text : string.Empty);
                    command.Parameters.AddWithValue("$D_A", (CheckBox_D_A.IsChecked ?? false) ? int.Parse(TextBox_D_A.Text) : 0);
                    command.Parameters.AddWithValue("$D_B", (CheckBox_D_B.IsChecked ?? false) ? int.Parse(TextBox_D_B.Text) : 0);
                    command.Parameters.AddWithValue("$phonenumber", TextBox_PhoneNumber.Text.Length > 0 ? TextBox_PhoneNumber.Text : string.Empty);
                    command.Parameters.AddWithValue("$datetime", select.Tick);
                    command.Parameters.AddWithValue("$sum", CalSum());
                    command.Parameters.AddWithValue("$target_name", TextBox_TargetName.Text.Length > 0 ? TextBox_TargetName.Text : string.Empty);
                    command.Parameters.AddWithValue("$employee_name", TextBox_Employee_Name.Text.Length > 0 ? TextBox_Employee_Name.Text : string.Empty);
                    command.Parameters.AddWithValue("$valid", 1);

                    int rowsAffected = command.ExecuteNonQuery();
                    if (rowsAffected > 0)
                    {
                        Console.WriteLine("Record updated successfully.");
                    }
                    else
                    {
                        Console.WriteLine("No record found with the specified ID.");
                    }
                }
                catch (Exception ex)
                {
                    MessageBoxResult result = MessageBox.Show(
                        "Error",
                        "Fail EditSelected",
                        MessageBoxButton.OK,
                        MessageBoxImage.Error
                    );

                }
            }
            ReadAllFromDb();
        }
        private void DeclineSelected()
        {
            //@"  UPDATE users
            //    SET user_name= $userName
            //    WHERE id = $id;";
            string DbFilePath = Path.Combine(DocumentPath, DataFolder, DbFileName);
            string _connectionString = "Data Source=" + DbFilePath;
            ListItem select = MainList.SelectedItem as ListItem;

            if (select == null)
                return;

            using (var connection = new SqliteConnection(_connectionString))
            {
                connection.Open();
                var command = connection.CreateCommand();
                command.CommandText = @" UPDATE userdata SET valid= $valid WHERE id == $id;";
                command.Parameters.AddWithValue("$valid", -1);
                command.Parameters.AddWithValue("$id", select.Id);
                //int id = Convert.ToInt32((object)command.ExecuteScalar());
                command.ExecuteNonQuery();
                //return id;
            }
            ReadAllFromDb();
        }
        private void ClearList()
        {
            ClearItems();
        }
        private void AddList(ListItem listItem  )
        {
            this.Dispatcher.Invoke(() => {
                Items.Add(listItem);
            });
        }
        private int Insert()
        {
            System.Diagnostics.Debug.WriteLine("Insert");
            string DbFilePath = Path.Combine(DocumentPath, DataFolder, DbFileName);
            string _connectionString = "Data Source=" + DbFilePath;
            using (var connection = new SqliteConnection(_connectionString))
            {
                connection.Open();
                var command = connection.CreateCommand();
                command.CommandText =
                    @"  INSERT INTO userdata (name ,zodiac ,age,S_A ,S_B ,S_C ,S_C_type ,S_D ,S_E ,address ,D_A ,D_B ,phonenumber ,datetime ,sum ,target_name ,employee_name ,valid) 
                        values ($name ,$zodiac ,$age ,$S_A ,$S_B ,$S_C ,$S_C_type ,$S_D ,$S_E ,$address ,$D_A ,$D_B ,$phonenumber ,$datetime ,$sum ,$target_name ,$employee_name ,$valid);
                        select last_insert_rowid();
                    ";
                command.Parameters.AddWithValue("$name",(TextBox_Name.Text.Length>0) ? TextBox_Name.Text:String.Empty);
                command.Parameters.AddWithValue("$zodiac", (TextBox_Zodiac.Text.Length > 0) ? TextBox_Zodiac.Text : String.Empty);
                command.Parameters.AddWithValue("$age", (TextBox_Age.Text.Length > 0) ? Int32.Parse(TextBox_Age.Text) : 0);
                command.Parameters.AddWithValue("$S_A", (bool)CheckBox_S_A.IsChecked  ? Int32.Parse(TextBox_S_A.Text) : 0);
                command.Parameters.AddWithValue("$S_B", (bool)CheckBox_S_B.IsChecked ? Int32.Parse(TextBox_S_B.Text): 0);
                command.Parameters.AddWithValue("$S_C", (bool)CheckBox_S_C.IsChecked ? Int32.Parse(TextBox_S_C.Text): 0);
                command.Parameters.AddWithValue("$S_C_type", (bool)CheckBox_S_C.IsChecked ? ListBox_S_C.SelectedIndex : -1);
                command.Parameters.AddWithValue("$S_D", (bool)CheckBox_S_D.IsChecked ? Int32.Parse(TextBox_S_D.Text): 0);
                command.Parameters.AddWithValue("$S_E", (bool)CheckBox_S_E.IsChecked ? Int32.Parse(TextBox_S_E.Text): 0);
                command.Parameters.AddWithValue("$address", (TextBox_Address.Text.Length > 0) ? TextBox_Address.Text : String.Empty);
                command.Parameters.AddWithValue("$D_A", (bool)CheckBox_D_A.IsChecked ? Int32.Parse(TextBox_D_A.Text) : 0);
                command.Parameters.AddWithValue("$D_B", (bool)CheckBox_D_B.IsChecked ? Int32.Parse(TextBox_D_B.Text) : 0);
                command.Parameters.AddWithValue("$phonenumber", (TextBox_PhoneNumber.Text.Length > 0) ? TextBox_PhoneNumber.Text : String.Empty);
                command.Parameters.AddWithValue("$datetime", DateTime.Now.Ticks);
                command.Parameters.AddWithValue("$sum", CalSum());
                command.Parameters.AddWithValue("$target_name", (TextBox_TargetName.Text.Length > 0) ? TextBox_TargetName.Text : String.Empty);
                command.Parameters.AddWithValue("$employee_name", (TextBox_Employee_Name.Text.Length > 0) ? TextBox_Employee_Name.Text : String.Empty);
                command.Parameters.AddWithValue("$valid", 1);
                if(TextBox_Employee_Name.Text.Length <= 0)
                {
                    return -1;
                }
                int id = Convert.ToInt32((object)command.ExecuteScalar());
                return id;
                //command.Parameters.AddWithValue("", );
            }
        }
        private int CalSum()
        {
            int sum = 0;
            sum += (bool)CheckBox_S_A.IsChecked ? Int32.Parse(TextBox_S_A.Text) : 0;
            sum += (bool)CheckBox_S_B.IsChecked ? Int32.Parse(TextBox_S_B.Text) : 0;
            sum += (bool)CheckBox_S_C.IsChecked ? Int32.Parse(TextBox_S_C.Text) : 0;
            sum += (bool)CheckBox_S_D.IsChecked ? Int32.Parse(TextBox_S_D.Text) : 0;
            sum += (bool)CheckBox_S_E.IsChecked ? Int32.Parse(TextBox_S_E.Text) : 0;
            sum += (bool)CheckBox_D_A.IsChecked ? Int32.Parse(TextBox_D_A.Text) : 0;
            sum += (bool)CheckBox_D_B.IsChecked ? Int32.Parse(TextBox_D_B.Text) : 0;
            return sum;
        }
        private bool ShowMessageBox(string Text, string Caption)
        {
            string messageBoxText = Text;
            string caption = Caption;
            MessageBoxButton button = MessageBoxButton.OK;
            MessageBoxImage icon = MessageBoxImage.Warning;
            MessageBoxResult result;

            result = MessageBox.Show(messageBoxText, caption, button, icon, MessageBoxResult.OK);
            if(result == MessageBoxResult.OK)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        private void ClearData(object sender, RoutedEventArgs e)
        {
            TextBox_Name.Text = string.Empty;
            TextBox_Zodiac.Text = string.Empty;
            TextBox_Age.Text = string.Empty;
            CheckBox_S_A.IsChecked = false;
            CheckBox_S_B.IsChecked = false;
            CheckBox_S_C.IsChecked = false;
            ListBox_S_C.SelectedIndex = -1;
            CheckBox_S_D.IsChecked = false;
            CheckBox_S_E.IsChecked = false;
            CheckBox_D_A.IsChecked = false;
            CheckBox_D_B.IsChecked = false;
            TextBox_PhoneNumber.Text = string.Empty;
            TextBox_TargetName.Text = string.Empty;
        }

        private int SCTypeToIndex(string i)
        {
           
            string[] zodiacs = new string[] { "喪門", "五鬼官符", "死符", "歲破", "白虎", "天狗", "病符" };
            return Array.IndexOf(zodiacs,i);
        }

        private void Sum_MouseDown(object sender, MouseButtonEventArgs e)
        {
            TextBlock_Sum.Text = CalSum().ToString();
        }

        private void CheckBox_D_A_Checked(object sender, RoutedEventArgs e)
        {

        }

        private void Search(object sender, RoutedEventArgs e)
        {
            ClearList();
            if(CheckBox_D_A.IsChecked == true)
            {
                //Get_D_A();// read and write files
            }
            if (CheckBox_D_B.IsChecked == true)
            {
                Get_D_B();
            }
            if (CheckBox_S_A.IsChecked == true)
            {
                Get_S_A();
            }
            if (CheckBox_S_B.IsChecked == true)
            {
                Get_S_B();
            }
            if (CheckBox_S_C.IsChecked == true)
            {
                Get_S_C();
            }
            if (CheckBox_S_D.IsChecked == true)
            {
                Get_S_D();
            }
            if (CheckBox_S_E.IsChecked == true)
            {
                Get_S_E();
            }
            if (TextBox_PhoneNumber.Text != string.Empty) {
                Get_pPhoneNumber();
            }

            //ReadAllFromDb();
        }
        private void Get_pPhoneNumber()
        {
            string DbFilePath = Path.Combine(DocumentPath, DataFolder, DbFileName);
            string OutputFilePath = Path.Combine(DocumentPath, DataFolder, "香油錢");
            string _connectionString = "Data Source=" + DbFilePath;
            using (var connection = new SqliteConnection(_connectionString))
            {
                connection.Open();
                var command = connection.CreateCommand();
                command.CommandText = @"  select * from userdata WHERE phonenumber = '"+ TextBox_PhoneNumber.Text + @"' AND valid > 0 ORDER BY id DESC";

                using (var reader = command.ExecuteReader())
                {
                    ClearList();
                    while (reader.Read())
                    {
                        object[] datas = new object[reader.FieldCount];
                        reader.GetValues(datas);
                        ListItem output = new ListItem(datas);
                        string path = Path.Combine(OutputFilePath, output.Datetime + "_" + output.Name + "_" + "香油錢.txt");
                        List<string> lines = new List<string>();
                        AddList(output);
                        //lines.Add("姓名: " + output.Name);
                        //lines.Add("金額: " + output.D_A);
                        //lines.Add("地址: " + output.Address);
                        //lines.Add("電話: " + output.PhoneNumber);
                        //string[] dataArray = new string[]
                        //{
                        //    output.Name,
                        //    output.PhoneNumber,
                        //    output.Target_Name,
                        //    "",
                        //    "",
                        //    "",
                        //    "",
                        //    "",
                        //    output.D_A,
                        //    "",
                        //    "",
                        //    "",
                        //    "",
                        //    output.Address,
                        //};
                        //CreateWordFile(dataArray, OutputFilePath, output.Datetime + "_" + output.Name + "_" + "香油錢.docx");
                        //Directory.CreateDirectory(OutputFilePath);
                        //File.WriteAllLines(path,lines);
                    }
                }
            }
        }
        /// <summary>
        /// ["name"] = strings[0],
        ///["phonenumber"] = strings[1],
        ///["target_name"] = strings[2],
        ///            ["zodiac"] = strings[3],
        ///            ["age"] = strings[4],
        ///            ["S_A"] = strings[5],
        ///            ["S_B"] = strings[6],
        ///            ["S_C"] = strings[7],
        ///            ["D_A"] = strings[8],
        ///            ["S_D"] = strings[9],
        ///            ["S_E"] = strings[10],
        ///            ["employee_name"] = strings[11],
        ///            ["sum"] = strings[12]
        ///            ["address"] = strings[13],
        ///
        /// </summary>
        /// <param name="strings"></param>
        /// <param name="ResultFolder"></param>
        private void CreateWordFile(Dictionary<string, string> stringDic, string ResultFolder, string filename, string wordTemplateName) 
        {
            string filepath = Path.Combine( AppDomain.CurrentDomain.BaseDirectory, wordTemplateName);
            var docxBytes = WordRender.GenerateDocx(File.ReadAllBytes(filepath),stringDic);
            File.WriteAllBytes(
                Path.Combine(ResultFolder, filename),
                docxBytes);
        }
        
        //string[] dataArray = new string[]
        //{
        //    "",
        //    "",
        //    "",
        //    "",
        //    "",
        //    "",
        //    "",
        //    "",
        //    "",
        //    "",
        //    "",
        //    "",
        //    "",
        //    "",
        //    "",
        //};
        private void Get_D_A()
        {
            string DbFilePath = Path.Combine(DocumentPath, DataFolder, DbFileName);
            string OutputFilePath = Path.Combine(DocumentPath, DataFolder,"名單");
            string _connectionString = "Data Source=" + DbFilePath;
            const int numPerPage = 8;
            using (var connection = new SqliteConnection(_connectionString))
            {
                connection.Open();
                var command = connection.CreateCommand();
                //get targetnames
                command.CommandText = @"  SELECT DISTINCT target_name FROM userdata WHERE valid > 0; ";
                List<string> TargetNames = new List<string>();
                using (var reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        object[] datas = new object[reader.FieldCount];
                        reader.GetValues(datas);
                        TargetNames.Add(datas[0].ToString());
                    }

                    foreach (string targetName in TargetNames)
                    {
                        using ( command = new SqliteCommand("SELECT * FROM userdata WHERE target_name = $TargetName AND valid > 0", connection))
                        {
                            command.Parameters.AddWithValue("$TargetName", targetName);

                            using (SqliteDataReader readerMember = command.ExecuteReader())
                            {
                                List<ListItem> membersRaw = new List<ListItem>();
                                while (readerMember.Read())
                                {
                                    object[] datas = new object[readerMember.FieldCount];
                                    readerMember.GetValues(datas);
                                    ListItem output = new ListItem(datas);
                                    membersRaw.Add(output);
                                }
                                // done getting member try generate word
                                //calculate in block and in lantern
                                List<ListItem> members = new List<ListItem>();
                                List<ListItem> lanterns = new List<ListItem>();
                                foreach (ListItem output in membersRaw)
                                {
                                    //no lantern
                                    if (output.S_D.Equals("0") && output.S_E.Equals("0"))
                                    {
                                        //add to members
                                        members.Add(output);
                                    }
                                    //lantern only
                                    else if((!output.S_D.Equals("0") || !output.S_E.Equals("0"))
                                        && output.S_A.Equals("0")
                                        && output.S_B.Equals("0")
                                        && output.S_C.Equals("0"))
                                    {
                                        //add lantern
                                        lanterns.Add(output);
                                    }
                                    
                                    //lantern and block
                                    else if ((!output.S_D.Equals("0") || !output.S_E.Equals("0"))
                                        &&( output.S_A.Equals("0")
                                        || output.S_B.Equals("0")
                                        || output.S_C.Equals("0")))
                                    {
                                        //add both
                                        members.Add(output);
                                        lanterns.Add(output);
                                    }
                                    else
                                    {
                                        MessageBox.Show("Exceptiont D_A");
                                    }
                                }

                                //calculate pages
                                int addLeft = 0;
                                if(members.Count % numPerPage > 0)
                                {
                                    //has left
                                    addLeft = 1;
                                }
                                int pages = (members.Count / numPerPage) + addLeft;
                                for(int i = 0; i < pages; i++)
                                {
                                    ListItem target = null;
                                    bool IsTargetGet = false;
                                    //get target info
                                    using (command = new SqliteCommand("SELECT * FROM userdata WHERE name = $TargetName AND valid > 0 ORDER BY datetime DESC", connection))
                                    {
                                        command.Parameters.AddWithValue("$TargetName", members[i * numPerPage + 0].Target_Name);
                                        using (SqliteDataReader readertarget = command.ExecuteReader()) 
                                        {
                                            if (readertarget.Read())
                                            {
                                                object[] datas = new object[readertarget.FieldCount];
                                                readertarget.GetValues(datas);
                                                target = new ListItem(datas);
                                                IsTargetGet = true;
                                            }
                                        }
                                    }

                                    var data = new Dictionary<string, string>
                                    {
                                        { "nameApply", IsTargetGet ? target.Name:"" },
                                        { "phoneNum", IsTargetGet ? target.PhoneNumber : "" },
                                        { "address", IsTargetGet ? target.Address : "" },
                                    };

                                    int memberCount = members.Count; // Get the total number of members
                                    int maxIndex = Math.Min(numPerPage, memberCount - i * numPerPage); // Calculate how many members are available starting at i * 10


                                    // Dynamically add entries for indices 0 to 9
                                    for (int n = 0; n < maxIndex; n++)
                                    {
                                        data.Add($"Name{n}", members[i * numPerPage + n].Name);
                                        data.Add($"zod{n}", members[i * numPerPage + n].Zodiac);
                                        data.Add($"age{n}", members[i * numPerPage + n].Age);
                                        data.Add($"1v{n}", ""); // Placeholder for additional values
                                        data.Add($"2v{n}", "");
                                        data.Add($"3v{n}", "");
                                        data.Add($"sin{n}", "");
                                    }

                                    // Fill the remaining keys with empty strings
                                    for (int n = maxIndex; n < numPerPage; n++)
                                    {
                                        data.Add($"Name{n}", "");
                                        data.Add($"zod{n}", "");
                                        data.Add($"age{n}", "");
                                        data.Add($"1v{n}", "");
                                        data.Add($"2v{n}", "");
                                        data.Add($"3v{n}", "");
                                        data.Add($"sin{n}", "");
                                    }

                                    string S_D_String = "";
                                    string S_E_String = "";
                                    //prepare lantern strings
                                    foreach (ListItem lant in lanterns)
                                    {
                                       
                                        if (!lant.S_D.Equals("0")) 
                                        {
                                            S_D_String = lant.Name + " ";
                                            
                                        }
                                        if (!lant.S_E.Equals("0")) 
                                        {
                                            S_E_String = lant.Name + " ";
                                           
                                        }
                                    }

                                    // Add the remaining static keys
                                    data.Add("suma", "");
                                    data.Add("sumb", "");
                                    data.Add("sumc", "");
                                    data.Add("sum", "");
                                    data.Add("oil", "");
                                    data.Add("bainame", S_D_String);
                                    data.Add("bsum", "");
                                    data.Add("naname", S_E_String);
                                    data.Add("nsum", "");
                                    data.Add("employname", "");
                                    data.Add("year", "");
                                    data.Add("mon", "");
                                    data.Add("day", ""); 
                                    data.Add("totalsum", "");
                                    Directory.CreateDirectory(OutputFilePath);
                                    //"shinword2.docx"
                                    CreateWordFile(data, OutputFilePath, members[i * numPerPage + 0].Target_Name + "_" + i.ToString() + "_" + "名單.docx", "shinWord3.docx");
                                    //File.WriteAllLines(path, lines);
                                }
                                
                            }
                        }
                    }


                   
                }
            }
        }



        private void Get_D_B()
        {
            string DbFilePath = Path.Combine(DocumentPath, DataFolder, DbFileName);
            string OutputFilePath = Path.Combine(DocumentPath, DataFolder, "磧爐錢");
            string _connectionString = "Data Source=" + DbFilePath;
            using (var connection = new SqliteConnection(_connectionString))
            {
                connection.Open();
                var command = connection.CreateCommand();
                command.CommandText = @"  select * from userdata WHERE D_B > 0 AND valid > 0 ORDER BY id DESC";

                using (var reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        object[] datas = new object[reader.FieldCount];
                        reader.GetValues(datas);
                        ListItem output = new ListItem(datas);
                        string path = Path.Combine(OutputFilePath, output.Datetime + "_" + output.Name + "_" + "磧爐錢.txt");
                        List<string> lines = new List<string>();
                        lines.Add("姓名: " + output.Name);
                        lines.Add("金額: " + output.D_B);
                        lines.Add("地址: " + output.Address);
                        lines.Add("電話: " + output.PhoneNumber);
                        Directory.CreateDirectory(OutputFilePath);
                        File.WriteAllLines(path, lines);
                    }
                }
            }
        }
        private void Get_S_A()
        {
            string DbFilePath = Path.Combine(DocumentPath, DataFolder, DbFileName);
            string OutputFilePath = Path.Combine(DocumentPath, DataFolder, "光明燈");
            string _connectionString = "Data Source=" + DbFilePath;
            using (var connection = new SqliteConnection(_connectionString))
            {
                connection.Open();
                var command = connection.CreateCommand();
                command.CommandText = @"  select * from userdata WHERE S_A > 0 AND valid > 0 ORDER BY id DESC";

                using (var reader = command.ExecuteReader())
                {
                    List<string> lines = new List<string>();
                    string path = Path.Combine(OutputFilePath, DateTime.Now.ToString("yyyy_MM_dd_HH_mm_ss") +  "_" + "光明燈.csv");
                    lines.Add("姓名,金額,電話,對象");
                    while (reader.Read())
                    {
                        object[] datas = new object[reader.FieldCount];
                        reader.GetValues(datas);
                        ListItem output = new ListItem(datas);
                        string row =string.Empty;
                        row +=  output.Name + ",";
                        row +=  output.S_A + ",";
                        row +=  output.PhoneNumber + ",";
                        row += output.Target_Name + ",";
                        lines.Add(row);
                    }
                    Directory.CreateDirectory(OutputFilePath);
                    File.WriteAllLines(path, lines);
                }
            }
        }
        private void Get_S_B()
        {
            string DbFilePath = Path.Combine(DocumentPath, DataFolder, DbFileName);
            string OutputFilePath = Path.Combine(DocumentPath, DataFolder, "太歲");
            string _connectionString = "Data Source=" + DbFilePath;
            using (var connection = new SqliteConnection(_connectionString))
            {
                connection.Open();
                var command = connection.CreateCommand();
                command.CommandText = @"  select * from userdata WHERE S_B > 0 AND valid > 0 ORDER BY id DESC";

                using (var reader = command.ExecuteReader())
                {
                    List<string> lines = new List<string>();
                    string path = Path.Combine(OutputFilePath, DateTime.Now.ToString("yyyy_MM_dd_HH_mm_ss") + "_" + "太歲.csv");
                    lines.Add("姓名,金額,電話,對象");
                    while (reader.Read())
                    {
                        object[] datas = new object[reader.FieldCount];
                        reader.GetValues(datas);
                        ListItem output = new ListItem(datas);
                        string row = string.Empty;
                        row += output.Name + ",";
                        row += output.S_B + ",";
                        row += output.PhoneNumber + ",";
                        row += output.Target_Name + ",";
                        lines.Add(row);
                    }
                    Directory.CreateDirectory(OutputFilePath);
                    File.WriteAllLines(path, lines);
                }
            }
        }
        private void Get_S_C()
        {
            string DbFilePath = Path.Combine(DocumentPath, DataFolder, DbFileName);
            string OutputFilePath = Path.Combine(DocumentPath, DataFolder, "制化");
            string _connectionString = "Data Source=" + DbFilePath;
            using (var connection = new SqliteConnection(_connectionString))
            {
                connection.Open();
                var command = connection.CreateCommand();
                command.CommandText = @"  select * from userdata WHERE S_C > 0 AND valid > 0 ORDER BY id DESC";

                using (var reader = command.ExecuteReader())
                {
                    List<string> lines = new List<string>();
                    string path = Path.Combine(OutputFilePath, DateTime.Now.ToString("yyyy_MM_dd_HH_mm_ss") + "_" + "制化.csv");
                    lines.Add("姓名,金額,電話,對象,種類");
                    while (reader.Read())
                    {
                        object[] datas = new object[reader.FieldCount];
                        reader.GetValues(datas);
                        ListItem output = new ListItem(datas);
                        string row = string.Empty;
                        row += output.Name + ",";
                        row += output.S_C + ",";
                        row += output.PhoneNumber + ",";
                        row += output.Target_Name + ",";
                        row += output.S_C_Type + ",";
                        lines.Add(row);
                    }
                    Directory.CreateDirectory(OutputFilePath);
                    File.WriteAllLines(path, lines);
                }
            }
        }
        private void Get_S_D()
        {
            string DbFilePath = Path.Combine(DocumentPath, DataFolder, DbFileName);
            string OutputFilePath = Path.Combine(DocumentPath, DataFolder, "拜亭燈籠");
            string _connectionString = "Data Source=" + DbFilePath;
            using (var connection = new SqliteConnection(_connectionString))
            {
                connection.Open();
                var command = connection.CreateCommand();
                command.CommandText = @"  select * from userdata WHERE S_D > 0 AND valid > 0 ORDER BY id DESC";

                using (var reader = command.ExecuteReader())
                {
                    List<string> lines = new List<string>();
                    string path = Path.Combine(OutputFilePath, DateTime.Now.ToString("yyyy_MM_dd_HH_mm_ss") + "_" + "拜亭燈籠.csv");
                    lines.Add("姓名,金額,電話,對象");
                    while (reader.Read())
                    {
                        object[] datas = new object[reader.FieldCount];
                        reader.GetValues(datas);
                        ListItem output = new ListItem(datas);
                        string row = string.Empty;
                        row += output.Name + ",";
                        row += output.S_D + ",";
                        row += output.PhoneNumber + ",";
                        row += output.Target_Name + ",";
                        lines.Add(row);
                    }
                    Directory.CreateDirectory(OutputFilePath);
                    File.WriteAllLines(path, lines);
                }
            }
        }
        private void Get_S_E()
        {
            string DbFilePath = Path.Combine(DocumentPath, DataFolder, DbFileName);
            string OutputFilePath = Path.Combine(DocumentPath, DataFolder, "內殿燈籠");
            string _connectionString = "Data Source=" + DbFilePath;
            using (var connection = new SqliteConnection(_connectionString))
            {
                connection.Open();
                var command = connection.CreateCommand();
                command.CommandText = @"  select * from userdata WHERE S_E > 0 AND valid > 0 ORDER BY id DESC";

                using (var reader = command.ExecuteReader())
                {
                    List<string> lines = new List<string>();
                    string path = Path.Combine(OutputFilePath, DateTime.Now.ToString("yyyy_MM_dd_HH_mm_ss") + "_" + "內殿燈籠.csv");
                    lines.Add("姓名,金額,電話,對象");
                    while (reader.Read())
                    {
                        object[] datas = new object[reader.FieldCount];
                        reader.GetValues(datas);
                        ListItem output = new ListItem(datas);
                        string row = string.Empty;
                        row += output.Name + ",";
                        row += output.S_E + ",";
                        row += output.PhoneNumber + ",";
                        row += output.Target_Name + ",";
                        lines.Add(row);
                    }
                    Directory.CreateDirectory(OutputFilePath);
                    File.WriteAllLines(path, lines);
                }
            }
        }

        private void MainList_Selected(object sender, SelectionChangedEventArgs e)
        {
            ListItem select = MainList.SelectedItem as ListItem;
            if (select != null)
            {
                TextBox_Name.Text = select.Name;
                TextBox_Zodiac.Text = select.Zodiac;
                TextBox_Age.Text = select.Age;
                CheckBox_S_A.IsChecked = int.Parse(select.S_A) > 0 ? true : false;
                CheckBox_S_B.IsChecked = int.Parse(select.S_B) > 0 ? true : false;
                CheckBox_S_C.IsChecked = int.Parse(select.S_C) > 0 ? true : false;
                ListBox_S_C.SelectedIndex = SCTypeToIndex(select.S_C_Type);
                CheckBox_S_D.IsChecked = int.Parse(select.S_D) > 0 ? true : false;
                CheckBox_S_E.IsChecked = int.Parse(select.S_E) > 0 ? true : false;
                CheckBox_D_A.IsChecked = int.Parse(select.D_A) > 0 ? true : false;
                CheckBox_D_B.IsChecked = int.Parse(select.D_B) > 0 ? true : false;
                TextBox_PhoneNumber.Text = select.PhoneNumber;
                TextBox_TargetName.Text = select.Target_Name;
                TextBox_Employee_Name.Text = select.Employee_Name;
                TextBox_Address.Text = select.Address;
            }
        }

        private void Decline_Click(object sender, RoutedEventArgs e)
        {
            DeclineSelected();
        }

        private void Edit_Click(object sender, RoutedEventArgs e)
        {
            EditSelected();
        }

        private void GenerateShinWord(object sender, RoutedEventArgs e)
        {
            Get_D_A();
        }
    }

    public class BoolToVisibilityConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            //bool IsCheck = (bool)value;
            //if(IsCheck)
            //    return Visibility.Visible;
            //return Visibility.Collapsed;
            if ((Visibility)value == Visibility.Visible)
                return true;
            return false;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            //if((Visibility)value == Visibility.Visible)
            //    return true;
            //return false;
            bool IsCheck = (bool)value;
            if (IsCheck)
                return Visibility.Visible;
            return Visibility.Collapsed;
        }
    }
    public class ListItem
    {
        public string? Id { get; set; }
        public string? Name { get; set; }
        public string? Zodiac { get; set; }
        public string? Age { get; set; }
        /// <summary>
        /// 光明燈
        /// </summary>
        public string? S_A { get; set; }//光明燈
        /// <summary>
        /// 太歲
        /// </summary>
        public string? S_B { get; set; }
        /// <summary>
        /// 制化
        /// </summary>
        public string? S_C { get; set; }
        /// <summary>
        /// 制化 type 0"喪門", 1"五鬼官符", 2"死符", 3"歲破", 4"白五" , 5"天狗", 6"病符"
        /// </summary>
        public string? S_C_Type { get; set; }
        /// <summary>
        /// 香油錢
        /// </summary>
        public string? D_A { get; set; }
        /// <summary>
        /// 磧爐錢
        /// </summary>
        public string? D_B { get; set; }
        /// <summary>
        /// 拜亭燈龍
        /// </summary>
        public string? S_D { get; set; }
        /// <summary>
        /// 內殿燈籠
        /// </summary>
        public string? S_E { get; set; }
        public string? PhoneNumber { get; set; }
        public string? Datetime { get; set; }
        public string? Sum { get; set; }
        public string? Target_Name { get; set; }
        public string? Address { get; set; }
        public string? Employee_Name { get; set; }
        public string? Valid { get; set; }
        public long Tick { get; set; }
        public ListItem(object[] datas) 
        {
            Id = datas[0].ToString();
            Name = datas[1].ToString();
            Zodiac = datas[2].ToString();
            Age = datas[3].ToString();
            S_A = datas[4].ToString();
            S_B = datas[5].ToString();
            S_C = datas[6].ToString();
            S_C_Type = IndexToSCType(int.Parse(datas[7].ToString()));
            D_A = datas[11].ToString();
            D_B = datas[12].ToString();
            S_D = datas[8].ToString();
            S_E = datas[9].ToString();
            PhoneNumber = datas[13].ToString();
            Datetime = DateTimeParse(datas);
            Sum = datas[15].ToString();
            Target_Name = datas[16].ToString();
            Address = datas[10].ToString();
            Employee_Name = datas[17].ToString();
            Valid = datas[18].ToString();
            Tick = DateTimeTickParse(datas);
        }

       
        private string IndexToSCType(int i)
        {
            if (i < 0)
            {
                return String.Empty;
            }
            string[] zodiacs = new string[] { "喪門", "五鬼官符", "死符", "歲破", "白虎" , "天狗", "病符" };
            return zodiacs[i];
        }
        private string DateTimeParse(object[] datas)
        {
            object data14 = datas[14]; // Assuming datas[14] can be either long or string.
            string datetime;

            if (data14 is long longValue)
            {
                // Handle as long
                datetime = new DateTime(longValue).ToString("yyyy_MM_dd_HH_mm_ss");
            }
            else if (data14 is string stringValue)
            {
                if (long.TryParse(stringValue, out long parsedLong))
                {
                    // If the string can be parsed as a long
                    datetime = new DateTime(parsedLong).ToString("yyyy_MM_dd_HH_mm_ss");
                }
                else
                {
                    // Handle invalid string (e.g., assign a default value or throw an error)
                    datetime = "Invalid_Date";
                }
            }
            else
            {
                // Handle unexpected types (e.g., assign a default value or throw an error)
                datetime = "Invalid_Date";
            }
            return datetime;
        }

        private long DateTimeTickParse(object[] datas)
        {
            object data14 = datas[14]; // Assuming datas[14] can be either long or string.
            long datetime;

            if (data14 is long longValue)
            {
                // Handle as long
                datetime = new DateTime(longValue).Ticks;
            }
            else if (data14 is string stringValue)
            {
                if (long.TryParse(stringValue, out long parsedLong))
                {
                    // If the string can be parsed as a long
                    datetime = new DateTime(parsedLong).Ticks;
                }
                else
                {
                    // Handle invalid string (e.g., assign a default value or throw an error)
                    datetime =DateTime.Now.Ticks;
                }
            }
            else
            {
                // Handle unexpected types (e.g., assign a default value or throw an error)
                datetime = DateTime.Now.Ticks;
            }
            return datetime;
        }
    }
    public class DigitBox : TextBox
    {
        #region Constructors
        /*
         * The default constructor
         */
        public DigitBox()
        {
            TextChanged += new TextChangedEventHandler(OnTextChanged);
            KeyDown += new KeyEventHandler(OnKeyDown);
        }
        #endregion

        #region Properties
        new public String Text
        {
            get { return base.Text; }
            set
            {
                base.Text = LeaveOnlyNumbers(value);
            }
        }

        #endregion

        #region Functions
        private bool IsNumberKey(Key inKey)
        {
            if (inKey < Key.D0 || inKey > Key.D9)
            {
                if (inKey < Key.NumPad0 || inKey > Key.NumPad9)
                {
                    return false;
                }
            }
            return true;
        }

        private bool IsDelOrBackspaceOrTabKey(Key inKey)
        {
            return inKey == Key.Delete || inKey == Key.Back || inKey == Key.Tab;
        }

        private string LeaveOnlyNumbers(String inString)
        {
            String tmp = inString;
            foreach (char c in inString.ToCharArray())
            {
                if (!System.Text.RegularExpressions.Regex.IsMatch(c.ToString(), "^[0-9]*$"))
                {
                    tmp = tmp.Replace(c.ToString(), "");
                }
            }
            return tmp;
        }
        #endregion

        #region Event Functions
        protected void OnKeyDown(object sender, KeyEventArgs e)
        {
            e.Handled = !IsNumberKey(e.Key) && !IsDelOrBackspaceOrTabKey(e.Key);
        }

        protected void OnTextChanged(object sender, TextChangedEventArgs e)
        {
            base.Text = LeaveOnlyNumbers(Text);
        }
        #endregion
    }

    public static class WordRender
    {
        static void ReplaceParserTag(this OpenXmlElement elem, Dictionary<string, string> data)
        {
            var pool = new List<Run>();
            var matchText = string.Empty;
            var hiliteRuns = elem.Descendants<Run>() //找出鮮明提示
                .Where(o => o.RunProperties?.Elements<Highlight>().Any() ?? false).ToList();

            foreach (var run in hiliteRuns)
            {
                var t = run.InnerText;
                if (t.StartsWith("["))
                {
                    pool = new List<Run> { run };
                    matchText = t;
                }
                else
                {
                    matchText += t;
                    pool.Add(run);
                }
                if (t.EndsWith("]"))
                {
                    var m = Regex.Match(matchText, @"\[\$(?<n>\w+)\$\]");
                    if (m.Success && data.ContainsKey(m.Groups["n"].Value))
                    {
                        var firstRun = pool.First();
                        firstRun.RemoveAllChildren<Text>();
                        firstRun.RunProperties.RemoveAllChildren<Highlight>();
                        var newText = data[m.Groups["n"].Value];
                        var firstLine = true;
                        foreach (var line in Regex.Split(newText, @"\\n"))
                        {
                            if (firstLine) firstLine = false;
                            else firstRun.Append(new Break());
                            firstRun.Append(new Text(line));
                        }
                        pool.Skip(1).ToList().ForEach(o => o.Remove());
                    }
                }

            }
        }
        public static byte[] GenerateDocx(byte[] template, Dictionary<string, string> data)
        {
            using (var ms = new MemoryStream())
            {
                ms.Write(template, 0, template.Length);
                using (var docx = WordprocessingDocument.Open(ms, true))
                {
                    docx.MainDocumentPart.HeaderParts.ToList().ForEach(hdr =>
                    {
                        hdr.Header.ReplaceParserTag(data);
                    });
                    docx.MainDocumentPart.FooterParts.ToList().ForEach(ftr =>
                    {
                        ftr.Footer.ReplaceParserTag(data);
                    });
                    docx.MainDocumentPart.Document.Body.ReplaceParserTag(data);
                    docx.Save();
                }
                return ms.ToArray();
            }
        }
    }
}
