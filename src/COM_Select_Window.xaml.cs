﻿using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Ports;
using System.Linq;
using System.Management;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using System.Windows.Threading;

namespace HP_3478A
{
    /// <summary>
    /// Interaction logic for COM_Select_Window.xaml
    /// </summary>
    public partial class COM_Select_Window : Window
    {
        //Codes for Info Log Color Palette
        int Success_Code = 0;
        int Error_Code = 1;
        int Warning_Code = 2;
        int Config_Code = 3;
        int Message_Code = 4;

        //List of COM Ports stored in this
        List<string> portList;

        //COM Port Information, updated by GUI
        string COM_Port_Name = "";
        int COM_BaudRate_Value = 115200;
        int COM_Parity_Value = 0;
        int COM_StopBits_Value = 1;
        int COM_DataBits_Value = 8;
        int COM_Handshake_Value = 0;
        int COM_WriteTimeout_Value = 3000;
        int COM_ReadTimeout_Value = 3000;
        bool COM_RtsEnable = false;
        int COM_GPIB_Address_Value = 1;

        //Save Data Directory
        string folder_Directory;

        public COM_Select_Window()
        {
            InitializeComponent();
            Get_COM_List();
            getSoftwarePath();
            insert_Log("Make sure GPIB Address is correct.", Message_Code);
            insert_Log("Choose the correct COM port from the list.", Message_Code);
            insert_Log("Try AR488 ++rst button to reset the Adpater.", Message_Code);
            insert_Log("Click the Connect button when you are ready.", Message_Code);
        }

        private void getSoftwarePath()
        {
            try
            {
                folder_Directory = System.IO.Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\" + "Log Data (HP3478A)";
                insert_Log("Test Data will be saved inside the software directory.", Config_Code);
                insert_Log(folder_Directory, Config_Code);
                insert_Log("Click the Select button to select another directory.", Config_Code);
            }
            catch (Exception)
            {
                insert_Log("Cannot get software directory path. Choose a new directory.", Error_Code);
            }
        }

        private int folderCreation(string folderPath)
        {
            try
            {
                Directory.CreateDirectory(folderPath);
                return (0);
            }
            catch (Exception)
            {
                insert_Log("Cannot create test data folder. Choose another file directory.", Error_Code);
                return (1);
            }
        }

        private bool COM_Config_Updater()
        {
            COM_Port_Name = COM_Port.Text.ToUpper().Trim();

            string BaudRate = COM_Bits.SelectedItem.ToString().Split(new string[] { ": " }, StringSplitOptions.None).Last();
            COM_BaudRate_Value = Int32.Parse(BaudRate);

            string DataBits = COM_DataBits.SelectedItem.ToString().Split(new string[] { ": " }, StringSplitOptions.None).Last();
            COM_DataBits_Value = Int32.Parse(DataBits);

            bool isNum = int.TryParse(COM_write_timeout.Text.Trim(), out int Value);
            if (isNum == true & Value > 0)
            {
                COM_WriteTimeout_Value = Value;
                COM_write_timeout.Text = Value.ToString();
            }
            else
            {
                COM_write_timeout.Text = "1000";
                insert_Log("Write Timeout must be a positive integer.", Error_Code);
                return false;
            }

            isNum = int.TryParse(COM_read_timeout.Text.Trim(), out Value);
            if (isNum == true & Value > 0)
            {
                COM_ReadTimeout_Value = Value;
                COM_read_timeout.Text = Value.ToString();
            }
            else
            {
                COM_read_timeout.Text = "1000";
                insert_Log("Read Timeout must be a positive integer.", Error_Code);
                return false;
            }

            isNum = int.TryParse(GPIB_Address.Text.Trim(), out Value);
            if (isNum == true & Value > 0)
            {
                COM_GPIB_Address_Value = Value;
                GPIB_Address.Text = Value.ToString();
            }
            else
            {
                GPIB_Address.Text = "1";
                insert_Log("GPIB Address must be a positive integer.", Error_Code);
                return false;
            }

            string Parity = COM_Parity.SelectedItem.ToString().Split(new string[] { ": " }, StringSplitOptions.None).Last();
            switch (Parity)
            {
                case "Even":
                    COM_Parity_Value = 2;
                    break;
                case "Odd":
                    COM_Parity_Value = 1;
                    break;
                case "None":
                    COM_Parity_Value = 0;
                    break;
                case "Mark":
                    COM_Parity_Value = 3;
                    break;
                case "Space":
                    COM_Parity_Value = 4;
                    break;
                default:
                    COM_Parity_Value = 0;
                    break;
            }

            string StopBits = COM_Stop.SelectedItem.ToString().Split(new string[] { ": " }, StringSplitOptions.None).Last();
            switch (StopBits)
            {
                case "1":
                    COM_StopBits_Value = 1;
                    break;
                case "1.5":
                    COM_StopBits_Value = 3;
                    break;
                case "2":
                    COM_StopBits_Value = 2;
                    break;
                default:
                    COM_StopBits_Value = 1;
                    break;
            }

            string Flow = COM_Flow.SelectedItem.ToString().Split(new string[] { ": " }, StringSplitOptions.None).Last();
            switch (Flow)
            {
                case "Xon/Xoff":
                    COM_Handshake_Value = 1;
                    break;
                case "Hardware":
                    COM_Handshake_Value = 2;
                    break;
                case "None":
                    COM_Handshake_Value = 0;
                    break;
                default:
                    COM_Handshake_Value = 1;
                    break;
            }

            string rts = COM_rtsEnable.SelectedItem.ToString().Split(new string[] { ": " }, StringSplitOptions.None).Last();
            switch (rts)
            {
                case "True":
                    COM_RtsEnable = true;
                    break;
                case "False":
                    COM_RtsEnable = false;
                    break;
                default:
                    COM_RtsEnable = true;
                    break;
            }

            return true;
        }

        private void insert_Log(string Message, int Code)
        {
            SolidColorBrush Color = Brushes.Black;
            string Status = "";
            if (Code == Error_Code) //Error Message
            {
                Status = "[Error]";
                Color = Brushes.Red;
            }
            else if (Code == Success_Code) //Success Message
            {
                Status = "[Success]";
                Color = Brushes.Green;
            }
            else if (Code == Warning_Code) //Warning Message
            {
                Status = "[Warning]";
                Color = Brushes.Orange;
            }
            else if (Code == Config_Code) //Config Message
            {
                Status = "";
                Color = Brushes.Blue;
            }
            else if (Code == Message_Code)//Standard Message
            {
                Status = "";
                Color = Brushes.Black;
            }
            this.Dispatcher.Invoke(DispatcherPriority.Background, new ThreadStart(delegate
            {
                Info_Log.Inlines.Add(new Run(Status + " " + Message + "\n") { Foreground = Color });
                Info_Scroll.ScrollToBottom();
            }));
        }

        private void Get_COM_List()
        {
            using (var searcher = new ManagementObjectSearcher("SELECT * FROM Win32_PnPEntity WHERE Caption like '%(COM%'"))
            {
                var portnames = SerialPort.GetPortNames();
                var ports = searcher.Get().Cast<ManagementBaseObject>().ToList().Select(p => p["Caption"].ToString());
                portList = portnames.Select(n => n + " - " + ports.FirstOrDefault(s => s.Contains('(' + n + ')'))).ToList();
                foreach (string p in portList)
                {
                    updateList(p);
                }
            }
        }

        private void updateList(string data)
        {
            ListBoxItem COM_itm = new ListBoxItem();
            COM_itm.Content = data;
            COM_List.Items.Add(COM_itm);
        }

        private void COM_Refresh_Click(object sender, RoutedEventArgs e)
        {
            COM_List.Items.Clear();
            Get_COM_List();
        }

        private void COM_List_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            try
            {
                string temp = COM_List.SelectedItem.ToString().Split(new string[] { ": " }, StringSplitOptions.None).Last();
                string COM = temp.Substring(0, temp.IndexOf(" -"));
                COM_Port.Text = COM;
                COM_Open_Check();

            }
            catch (Exception)
            {
                insert_Log("Select a Valid COM Port.", Warning_Code);
            }
        }

        private bool COM_Open_Check()
        {
            try
            {
                using (var sp = new SerialPort(COM_Port.Text, 9600, System.IO.Ports.Parity.None, 8, System.IO.Ports.StopBits.One))
                {
                    sp.WriteTimeout = 500;
                    sp.ReadTimeout = 500;
                    sp.Handshake = Handshake.None;
                    sp.RtsEnable = true;
                    sp.Open();
                    System.Threading.Thread.Sleep(100);
                    sp.Close();
                    insert_Log(COM_Port.Text + " is open and ready for communication.", Success_Code);
                }
            }
            catch (Exception Ex)
            {
                COM_Port.Text = string.Empty;
                insert_Log(Ex.ToString(), Error_Code);
                insert_Log(COM_Port.Text + " is closed. Probably being used by a software.", Error_Code);
                insert_Log("Try another COM Port or check if COM is already used by another software.", Message_Code);
                return false;
            }
            return true;
        }

        private bool Set_COM_Open_Check()
        {
            try
            {
                using (var serial = new SerialPort(COM_Port_Name, COM_BaudRate_Value, (Parity)COM_Parity_Value, COM_DataBits_Value, (StopBits)COM_StopBits_Value))
                {
                    serial.WriteTimeout = COM_WriteTimeout_Value;
                    serial.ReadTimeout = COM_ReadTimeout_Value;
                    serial.RtsEnable = COM_RtsEnable;
                    serial.Handshake = (Handshake)COM_Handshake_Value;
                    serial.Open();
                    System.Threading.Thread.Sleep(100);
                    serial.Close();
                    return true;
                }
            }
            catch (Exception)
            {
                COM_Port.Text = string.Empty;
                insert_Log(COM_Port.Text + " is closed. Probably being used by a software.", Error_Code);
                insert_Log("Try another COM Port or check if com is already used by another software.", Message_Code);
            }
            return true;
        }

        private (bool, string) Serial_Query(string command)
        {
            try
            {
                using (var serial = new SerialPort(COM_Port_Name, COM_BaudRate_Value, (Parity)COM_Parity_Value, COM_DataBits_Value, (StopBits)COM_StopBits_Value))
                {
                    serial.WriteTimeout = COM_WriteTimeout_Value;
                    serial.ReadTimeout = COM_ReadTimeout_Value;
                    serial.RtsEnable = COM_RtsEnable;
                    serial.Handshake = (Handshake)COM_Handshake_Value;
                    serial.Open();
                    serial.WriteLine("++addr " + COM_GPIB_Address_Value);
                    System.Threading.Thread.Sleep(100);
                    serial.WriteLine(command);
                    string data = serial.ReadLine();
                    serial.Close();
                    return (true, data);
                }
            }
            catch (Exception)
            {
                insert_Log("Serial Query Failed, check COM settings or connection.", Error_Code);
                return (false, "");
            }
        }

        private bool Serial_Write(string command)
        {
            try
            {
                using (var serial = new SerialPort(COM_Port_Name, COM_BaudRate_Value, (Parity)COM_Parity_Value, COM_DataBits_Value, (StopBits)COM_StopBits_Value))
                {
                    serial.WriteTimeout = COM_WriteTimeout_Value;
                    serial.ReadTimeout = COM_ReadTimeout_Value;
                    serial.RtsEnable = COM_RtsEnable;
                    serial.Handshake = (Handshake)COM_Handshake_Value;
                    serial.Open();
                    System.Threading.Thread.Sleep(100);
                    serial.WriteLine(command);
                    serial.Close();
                    return true;
                }
            }
            catch (Exception)
            {
                insert_Log("Serial Write Failed, check COM settings or connection.", Error_Code);
                return false;
            }
        }

        private void Select_Directory_Click(object sender, RoutedEventArgs e)
        {
            var Choose_Directory = new Ookii.Dialogs.Wpf.VistaFolderBrowserDialog();
            if (Choose_Directory.ShowDialog() == true)
            {
                folder_Directory = Choose_Directory.SelectedPath + @"\" + "Log Data (HP3478A)";
            }
            insert_Log("Test Data will be saved here: " + folder_Directory, Config_Code);
        }

        private void AR488_Version_Click(object sender, RoutedEventArgs e)
        {
            if (COM_Config_Updater() == true)
            {
                (bool check, string return_data) = Serial_Query("++ver");
                if (check == true)
                {
                    insert_Log(return_data, Success_Code);
                }
            }
            else
            {
                insert_Log("COM Info is invalid. Correct any errors and try again.", Error_Code);
            }
        }

        private void AR488_Reset_Click(object sender, RoutedEventArgs e)
        {
            if (COM_Config_Updater() == true)
            {
                if (Serial_Write("++rst") == true)
                {
                    insert_Log("Reset command was send successfully.", Success_Code);
                }
            }
        }

        private void AR488_GPIB_Address_Click(object sender, RoutedEventArgs e)
        {
            if (COM_Config_Updater() == true)
            {
                (bool check, string return_data) = Serial_Query("++addr");
                if (check == true)
                {
                    insert_Log("Current GPIB Address: " + return_data, Success_Code);
                }
            }
            else
            {
                insert_Log("COM Info is invalid. Correct any errors and try again.", Error_Code);
            }
        }

        private void Verify_3478A_Click(object sender, RoutedEventArgs e)
        {
            if (COM_Config_Updater() == true)
            {
                (bool check, string return_data) = Serial_Query("++read");
                if (check == true)
                {
                    insert_Log(return_data, Success_Code);
                    if (return_data.Trim().Length == 11)
                    {
                        insert_Log("Verify Successful.", Success_Code);
                    }
                    else
                    {
                        insert_Log("Verify Failed. Expected measurement value message with length 11.", Error_Code);
                        insert_Log("Try Again.", Error_Code);
                    }
                }
            }
            else
            {
                insert_Log("COM Info is invalid. Correct any errors and try again.", Error_Code);
            }
        }

        private bool Connect_verify_3478A()
        {
            try
            {
                using (var serial = new SerialPort(COM_Port_Name, COM_BaudRate_Value, (Parity)COM_Parity_Value, COM_DataBits_Value, (StopBits)COM_StopBits_Value))
                {
                    serial.WriteTimeout = COM_WriteTimeout_Value;
                    serial.ReadTimeout = COM_ReadTimeout_Value;
                    serial.RtsEnable = COM_RtsEnable;
                    serial.Handshake = (Handshake)COM_Handshake_Value;
                    serial.Open();
                    serial.WriteLine("++read");
                    string data = serial.ReadLine();
                    serial.Close();
                    if (data.Trim().Length == 11)
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                }
            }
            catch (Exception)
            {
                insert_Log("Serial Query Failed, check COM settings or connection.", Error_Code);
                return false;
            }
        }

        private void Connect_Click(object sender, RoutedEventArgs e)
        {
            if (folderCreation(folder_Directory) == 0)
            {
                folderCreation(folder_Directory + @"\" + "VDC");
                folderCreation(folder_Directory + @"\" + "VAC");
                folderCreation(folder_Directory + @"\" + "2WireOhms");
                folderCreation(folder_Directory + @"\" + "4WireOhms");
                folderCreation(folder_Directory + @"\" + "ADC");
                folderCreation(folder_Directory + @"\" + "AAC");
                if (COM_Config_Updater() == true)
                {
                    if (Set_COM_Open_Check() == true)
                    {
                        if (Connect_verify_3478A() == true)
                        {
                            if (Serial_Write("F1RAN5T1Z1D1") == true)
                            {
                                insert_Log("Connect Successful. Please wait for software initialization.", Success_Code);
                                Info_Log.Text = String.Empty;
                                Info_Log.Inlines.Clear();
                                Data_Updater();
                                this.Close();
                            }
                        }
                        else
                        {
                            insert_Log("Verify Failed. Try Again.", Error_Code);
                        }
                    }
                    else
                    {
                        insert_Log("COM Port is not open. Check if COM Port is in use.", Error_Code);
                        insert_Log("Connect Failed.", Error_Code);
                    }
                }
                else
                {
                    insert_Log("COM Info is invalid. Correct any errors and try again.", Error_Code);
                    insert_Log("Connect Failed.", Error_Code);
                }
            }
            else
            {
                insert_Log("Log Data Directory cannot be created on the selected path.", Error_Code);
                insert_Log("Choose another path by clicking the select button.", Error_Code);
            }
        }

        private void Data_Updater()
        {
            Serial_COM_Info.COM_Port = COM_Port_Name;
            Serial_COM_Info.COM_BaudRate = COM_BaudRate_Value;
            Serial_COM_Info.COM_Parity = COM_Parity_Value;
            Serial_COM_Info.COM_StopBits = COM_StopBits_Value;
            Serial_COM_Info.COM_DataBits = COM_DataBits_Value;
            Serial_COM_Info.COM_Handshake = COM_Handshake_Value;
            Serial_COM_Info.COM_WriteTimeout = COM_WriteTimeout_Value;
            Serial_COM_Info.COM_ReadTimeout = COM_ReadTimeout_Value;
            Serial_COM_Info.COM_RtsEnable = COM_RtsEnable;
            Serial_COM_Info.GPIB_Address = COM_GPIB_Address_Value;
            Serial_COM_Info.folder_Directory = folder_Directory;
            Serial_COM_Info.isConnected = true;
        }

        private void Info_Clear_Click(object sender, RoutedEventArgs e)
        {
            try 
            {
                Info_Log.Inlines.Clear();
                Info_Log.Text = string.Empty;
            } 
            catch (Exception) 
            {
                
            }
        }
    }
}
