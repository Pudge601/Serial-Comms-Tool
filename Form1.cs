using System;
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
            ComPort.Close();
            btnConnect.Text = "Connect";
            btnSend.Enabled = false;
            btnFileSend.Enabled = false;
            groupBox1.Enabled = true;
            
            }
              //whenever the connect button is clicked, it will check if the port is already open, call the disconnect function.
              // if the port is closed, call the connect function.
        private void btnConnect_Click(object sender, EventArgs e)
                                  
            {
            if (ComPort.IsOpen)
                {
                disconnect();
                }
            else
                {
                connect();
                }
            }

        private void btnClear_Click(object sender, EventArgs e)
            {
            //Clear the screen
            rtxtDataArea.Clear();
            txtSend.Clear();
            }
        // Function to send data to the serial port
        private void sendData(string writeData)
            {
            bool error = false;
            if (rdText.Checked == true)        //if text mode is selected, send data as tex
                {
                // Send the user's text straight out the port 
                ComPort.Write(writeData);
               
                // Show in the terminal window 
                rtxtDataArea.ForeColor = Color.Green;    //write sent text data in green colour              
                txtSend.Clear();                       //clear screen after sending data

                }
            else                    //if Hex mode is selected, send data in hexadecimal
                {
                try
                    {
                    // Convert the user's string of hex digits (example: E1 FF 1B) to a byte array
                    byte[] data = HexStringToByteArray(writeData);

                    // Send the binary data out the port
                  ComPort.Write(data, 0, data.Length);

                    // Show the hex digits on in the terminal window
                  rtxtDataArea.SelectionColor = Color.Blue;   //write Hex data in Blue
                  rtxtDataArea.AppendText(writeData.ToUpper() + "\n");
                    }
                catch (FormatException) { error = true; }
                    
                    // Inform the user if the hex string was not properly formatted
                    catch (ArgumentException) { error = true; }

                if (error) MessageBox.Show(this, "Not properly formatted hex string: " + writeData + "\n" + "example: E1 FF 1B", "Format Error", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                                                      
                }
            return;
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
        
        private void btnSend_Click(object sender, EventArgs e)
            {
            sendData(txtSend.Text);
            }
            //This event will be raised when the form is closing.
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
            {
            if (ComPort.IsOpen) ComPort.Close();  //close the port if open when exiting the application.
            }

        //Data recived from the serial port is coming from another thread context than the UI thread.
        //Instead of reading the content directly in the SerialPortDataReceived, we need to use a delegate.
        delegate void SetTextCallback(string text);
        private void SetText(string text)
        {
            //invokeRequired required compares the thread ID of the calling thread to the thread of the creating thread.
            // if these threads are different, it returns true
            if (this.rtxtDataArea.InvokeRequired)
            {
                rtxtDataArea.SelectionColor = Color.Green;    //write text data in Green colour
                //Debug.WriteLine("Received 2");

                SetTextCallback d = new SetTextCallback(SetText);              
                this.Invoke(d, new object[] { text });
            }
            else
            {
                //Debug.WriteLine("Received 1");
                this.rtxtDataArea.AppendText(text ); 
            }
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
            return Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog
            {
                InitialDirectory = MyDirectory(),
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
                hex.Append(Environment.NewLine);
            return hex.ToString();
        }


        private void btnFileSend_Click(object sender, EventArgs e)
        {
            var mre = new AutoResetEvent(false);
            var buffer = new StringBuilder();

            //Receive
            ComPort.DataReceived += (s, f) => //Think this is why we receive another empty line every time the send button is clicked
            {
                Debug.Print("Data Received");
                rtxtDataArea.SelectionColor = Color.Green;
                //Debug.Print("Data Received:");
                int bytestoreada = ComPort.BytesToRead; //reads 6 bytes here
                string bytestoreadastring = " " + bytestoreada + " ";
                //Debug.Print("bytestoreadastring" + bytestoreadastring + ".");
                byte[] dataa = new byte[bytestoreada]; // buffer to contain results of the read.
                ComPort.Read(dataa, 0, bytestoreada); //dump com port bytes to dataa
                string dataalengthstring = " " + dataa.Length + " ";
                //Debug.Print("dataalength: " + dataalengthstring + ".");
                string data = ByteArrayToString(dataa);
                if (data.Length > 0)
                {
                    this.rtxtDataArea.AppendText(data.ToUpper());
                }
                else
                {
                    rtxtDataArea.SelectionColor = Color.Red;
                    this.rtxtDataArea.AppendText("Blank Messsage Received!");
                }
                    //Debug.WriteLine("mre Set");
                buffer.Clear();
                mre.Set(); //allow loop to continue
            };

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
            var responseTimeout = TimeSpan.FromSeconds(1);

            foreach (var command in lines)
            {
                try
                {
                    byte[] data = HexStringToByteArray(command);
                    ComPort.Write(data, 0, data.Length);
                    Debug.Print("Data Sent");
                    string txData = string.Join(" ", data);
                    txData = (txData + System.Environment.NewLine);
                    rtxtDataArea.SelectionColor = Color.Blue;
                    this.rtxtDataArea.AppendText(txData.ToUpper());
                    
                    //this is where we block
                    if (!mre.WaitOne(responseTimeout))
                    {
                        Debug.WriteLine("Did not receive response");
                        //do something
                    }
                }
                catch (TimeoutException)
                {
                    Debug.WriteLine("Write took longer than expected");
                }
                catch
                {
                    Debug.WriteLine("Failed to write to port");
                }
            }

            Console.ReadLine();


        }

        private void rtxtDataArea_TextChanged(object sender, EventArgs e)
        {
            // set the current caret position to the end
                rtxtDataArea.SelectionStart = rtxtDataArea.Text.Length;
             // scroll it automatically
                rtxtDataArea.ScrollToCaret();
        }
    }
}
