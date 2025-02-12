﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO.Ports;  //This is a namespace that contains the SerialPort class
using System.Diagnostics;
using System.Reflection;
using System.IO;
using System.Threading;

namespace Serial_Communication
    {
    public partial class Form1 : Form
        {
        static EventWaitHandle mre = new AutoResetEvent(false);
        public int expectedResponseLength = 0;
        public int receivedLength = 0;
        public bool formConnectTrue = false;
        public string data = "";

        public Form1()
            {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e) 
            {
            updatePorts();           //Call this function everytime the page load 
                                     //to update port names
            CheckForIllegalCrossThreadCalls = false;
            cmbBaudRate.SelectedIndex = 5;
            cmbParity.SelectedIndex = 0;
            cmbDataBits.SelectedIndex = 1;
            cmbStopBits.SelectedIndex = 1;
            ComPort.DataReceived += new SerialDataReceivedEventHandler(onDataReceived);
        }
        private void updatePorts()
            {
            // Retrieve the list of all COM ports on your Computer
            string[] ports = SerialPort.GetPortNames();
            foreach (string port in ports)
                {
                cmbPortName.Items.Add(port);
                }
            }
        private SerialPort ComPort = new SerialPort();  //Initialise ComPort Variable as SerialPort
        private void connect()
            {
            Debug.WriteLine("Connect");
            bool error = false;

            // Check if all settings have been selected

            if (cmbPortName.SelectedIndex != -1 & cmbBaudRate.SelectedIndex != -1 & cmbParity.SelectedIndex != -1 & cmbDataBits.SelectedIndex != -1 & cmbStopBits.SelectedIndex != -1)
                {
                    //if yes than Set The Port's settings
                    ComPort.PortName = cmbPortName.Text;
                    ComPort.BaudRate = int.Parse(cmbBaudRate.Text);      //convert Text to Integer
                    ComPort.Parity = (Parity)Enum.Parse(typeof(Parity), cmbParity.Text); //convert Text to Parity
                    ComPort.DataBits = int.Parse(cmbDataBits.Text);        //convert Text to Integer
                    ComPort.StopBits = (StopBits)Enum.Parse(typeof(StopBits), cmbStopBits.Text);  //convert Text to stop bits
                    
                try  //always try to use this try and catch method to open your port. 
                     //if there is an error your program will not display a message instead of freezing.
                    {
                    //Open Port
                    ComPort.Open();
                    formConnectTrue = true;

                }
                catch (UnauthorizedAccessException) { error = true; }
                catch (System.IO.IOException) { error = true; }
                catch (ArgumentException) { error = true; }

                if (error) MessageBox.Show(this, "Could not open the COM port. Most likely it is already in use, has been removed, or is unavailable.", "COM Port unavailable", MessageBoxButtons.OK, MessageBoxIcon.Stop);

                }
            else
                {
                MessageBox.Show("Please select all the COM Serial Port Settings", "Serial Port Interface", MessageBoxButtons.OK, MessageBoxIcon.Stop);

                }
               //if the port is open, Change the Connect button to disconnect, enable the send button.
               //and disable the groupBox to prevent changing configuration of an open port.
            if (ComPort.IsOpen)
                {
                btnConnect.Text = "Disconnect";
                btnSend.Enabled = true;
                btnFileSend.Enabled = true;
                formConnectTrue = true;
                if (!rdText.Checked & !rdHex.Checked)  //if no data mode is selected, then select text mode by default
                    {
                    rdText.Checked = true;
                    }                
                groupBox1.Enabled = false;
                

                }
        }
              // Call this function to close the port.
        private void disconnect()
            {
            if (ComPort.IsOpen)
            {
                ComPort.Close();
            }
            btnConnect.Text = "Connect";
            btnSend.Enabled = false;
            btnFileSend.Enabled = false;
            groupBox1.Enabled = true;
            formConnectTrue = false;


            }
              //whenever the connect button is clicked, it will check if the port is already open, call the disconnect function.
              // if the port is closed, call the connect function.
        private void btnConnect_Click(object sender, EventArgs e)
                                  
            {
            if (formConnectTrue == true)
                {
                disconnect();
                }
            else
                {
                connect();
                }
            }

        //Clear tx/rx window
        private void btnClear_Click(object sender, EventArgs e)
            {
            //Clear the screen
            rtxtDataArea.Clear();
            }


               //Convert a string of hex digits (example: E1 FF 1B) to a byte array. 
               //The string containing the hex digits (with or without spaces)
              //Returns an array of bytes. </returns>
        private byte[] HexStringToByteArray(string s)
            {
            s = s.Replace(" ", "");
            byte[] buffer = new byte[s.Length / 2];
            for (int i = 0; i < s.Length; i += 2)
                buffer[i / 2] = (byte)Convert.ToByte(s.Substring(i, 2), 16);
            return buffer;
            }

        private void rtxtDataArea_TextChanged(object sender, EventArgs e) //Makes sure results window scrolls to the bottom of the page
        {
            // set the current caret position to the end
            rtxtDataArea.SelectionStart = rtxtDataArea.Text.Length;
            // scroll it automatically
            rtxtDataArea.ScrollToCaret();
        }

            //This event will be raised when the form is closing.
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
            {
            if (ComPort.IsOpen) ComPort.Close();  //close the port if open when exiting the application.
            }
        private void btnRefresh_Click(object sender, EventArgs e)
        {
            updatePorts();           //Call this function to update port names
        }

        private void cmbBaudRate_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

        }
        private string MyDirectory()
        {
            return Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location); //Find local .exe directory. Used when the file open dialog is called
        }

        private void button2_Click(object sender, EventArgs e) //Open File
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog
            {
                InitialDirectory = MyDirectory(), //Uses local .exe directory
                Title = "Browse Text Files",

                CheckFileExists = true,
                CheckPathExists = true,

                DefaultExt = "txt",
                Filter = "txt files (*.txt)|*.txt",
                FilterIndex = 2,
                RestoreDirectory = true,

                ReadOnlyChecked = true,
                ShowReadOnly = true
            };

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = openFileDialog1.FileName;
            }
        }

        public static string ByteArrayToString(byte[] ba)
        {
            StringBuilder hex = new StringBuilder(ba.Length * 2);
            foreach (byte b in ba)
                hex.AppendFormat("{0:x2} ", b);
            return hex.ToString();
        }

        
        // Function to send data to the serial port
        private void sendData(string writeData, System.TimeSpan responseTimeout)
        {
            try
            {
                byte[] data = HexStringToByteArray(writeData);
                expectedResponseLength = data.Length;
                Debug.WriteLine("Expected Length: " + expectedResponseLength);
                ComPort.Write(data, 0, data.Length);
                Debug.Print("Data Sent");
                string txData = string.Join(" ", writeData);
                txData = (txData + System.Environment.NewLine);
                rtxtDataArea.SelectionColor = Color.Blue;
                this.rtxtDataArea.AppendText(txData.ToUpper());

                //this is where we block. Wait for mre to be set in onDataReceived
                //mre.WaitOne(responseTimeout);
                if (!mre.WaitOne(responseTimeout))
                {
                    string noDataReceived = ("Did not receive response" + System.Environment.NewLine);
                    rtxtDataArea.SelectionColor = Color.Red;
                    this.rtxtDataArea.AppendText(noDataReceived.ToUpper());
                }
                else {
                    Debug.Print("mre received in send");
                }
                Debug.Print("was mre received in send?");

            }
            catch (TimeoutException)
            {
                string writeTimeout = ("Write took longer than expected" + System.Environment.NewLine);
                rtxtDataArea.SelectionColor = Color.Red;
                this.rtxtDataArea.AppendText(writeTimeout.ToUpper());
            }
            catch
            {
                string writeFail = ("Failed to write to port" + System.Environment.NewLine);
                rtxtDataArea.SelectionColor = Color.Red;
                this.rtxtDataArea.AppendText(writeFail.ToUpper());
            }
            return;
        }
        

        //Send one line from text box
        private void btnSend_Click(object sender, EventArgs e)
        {
            var singleResponseTimeout = TimeSpan.FromSeconds(1);
            sendData(txtSend.Text, singleResponseTimeout);
        }


        //Send all lines from file
        private void btnFileSend_Click(object sender, EventArgs e) 
        {
            //Send
            string[] lines = new String[0];
            try
            {
                lines = System.IO.File.ReadAllLines(@textBox1.Text);
            }
            catch
            {
                MessageBox.Show(this, "No File Open", "File Error", MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }
            var multiResponseTimeout = TimeSpan.FromSeconds(1);

            foreach (var command in lines)
            {
                sendData(command, multiResponseTimeout);
            }
        }






        //Data Received
        private void onDataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            Debug.Print("Data Received");
            rtxtDataArea.SelectionColor = Color.Green;
            int bytestoreada = ComPort.BytesToRead; //reads 6 bytes here
            receivedLength = receivedLength + bytestoreada;

            Debug.WriteLine("Bytes Received this time " + bytestoreada + ". Bytes received since last mre " + receivedLength + " .(Expected Response Length = " + expectedResponseLength + ")");
            byte[] dataa = new byte[bytestoreada]; // buffer to contain results of the read.
            ComPort.Read(dataa, 0, bytestoreada); //dump com port bytes to dataa
            if (bytestoreada < expectedResponseLength)
            {
                Debug.WriteLine("Concatenating Data");
                data = data + ByteArrayToString(dataa);
            }
            else
            {
                data = ByteArrayToString(dataa);
            }
            Debug.WriteLine("data.length " + data.Length);
            if (receivedLength < (expectedResponseLength))
            {
                rtxtDataArea.SelectionColor = Color.Red;
                Debug.WriteLine("Short Messsage Received!");
            }
            else
            {
                Debug.WriteLine("Print Data (" + data.Length + "bytes)");
                data = (data + Environment.NewLine);
                this.rtxtDataArea.AppendText(data.ToUpper());
                receivedLength = 0;
                data = "";
                mre.Set(); //allow loop to continue
                Debug.WriteLine("mre Set");
            }
        }
    }
}
