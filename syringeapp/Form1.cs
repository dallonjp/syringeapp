using System;
using System.Drawing;
using System.IO.Ports;
using System.Windows.Forms;
using System.Threading;
using System.Text;
using System.Globalization;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        public SerialPort port1;
        public static string handshake;
        public static bool isPushed = false;
        public static bool isPushedasp = false;
        public static bool isPushedinj = false;
        public static bool connected = false;
        public static bool ishomed = false;
        public static bool customaccel = false;
        public static bool aspirated = false;
        public static bool injected = false;

        public static char[] switchcode = { 'a', 'b', 'c', 'd', 'e' };

        public static double mmperstep = Properties.Settings.Default.pitch / ((360/Properties.Settings.Default.degree) * Properties.Settings.Default.microstp);
        public static double mlperstep=-1;
        public static double rate;
        public static double outputrate;
        public static double outputposition;
        public static double position=0;
        public static double vlm;
        public static double customsyrngevlm;
        public static double maxsteps=-1;
        public static double crntvlm = 0;

        public static string contrt;
        public static string contpsn;
        public static string sendit;
        public static string usrinput;
        public static string accelsend;
        System.Windows.Forms.Timer timer = new System.Windows.Forms.Timer();
        System.Diagnostics.Stopwatch watch = new System.Diagnostics.Stopwatch();
        public Form1()
        {
            InitializeComponent();
            //
            label1.Hide();
            label6.Hide();
            label12.Hide();
            label16.Hide();
            defaultStepsecToolStripMenuItem.Text = Properties.Settings.Default.accel.ToString();
            rotationPerStepdegreesToolStripMenuItem.Text = "Stepper degree: "+Properties.Settings.Default.degree.ToString();
            
            radioButton1.Text = Properties.Settings.Default.defaultml1.ToString() + " mL";
            radioButton2.Text= Properties.Settings.Default.defaultml2.ToString() + " mL";
            radioButton3.Text= Properties.Settings.Default.defaultml3.ToString()+" mL";
            mLToolStripMenuItem.Text= Properties.Settings.Default.defaultml1.ToString() + " mL";
            mLToolStripMenuItem1.Text = Properties.Settings.Default.defaultml2.ToString() + " mL";
            mLToolStripMenuItem2.Text = Properties.Settings.Default.defaultml3.ToString() + " mL";
            mLToolStripMenuItem3.Text = Properties.Settings.Default.defaultml1.ToString() + " mL";
            mLToolStripMenuItem4.Text = Properties.Settings.Default.defaultml2.ToString() + " mL";
            mLToolStripMenuItem5.Text = Properties.Settings.Default.defaultml3.ToString() + " mL";
            pitchToolStripMenuItem.Text = "Ball screw pitch= " + Properties.Settings.Default.pitch.ToString() + " mm";
            innerDiameterToolStripMenuItem.Text = "Inner Diamter= " + Properties.Settings.Default.defaultid1.ToString() + " mm";
            innerDiameterToolStripMenuItem1.Text= "Inner Diamter= " + Properties.Settings.Default.defaultid2.ToString() + " mm";
            innerDiameterToolStripMenuItem2.Text= "Inner Diamter= " + Properties.Settings.Default.defaultid3.ToString() + " mm";
            label5.Text="Disconnected";
            defaultStepsecToolStripMenuItem.Text = Properties.Settings.Default.accel.ToString();
            label16.Text = crntvlm.ToString();
            switch (Properties.Settings.Default.microstp)
            {
                case 1:
                    
                    fullToolStripMenuItem.BackColor = System.Drawing.Color.LightGray;
                    halfToolStripMenuItem.BackColor = default(Color);
                    quarterToolStripMenuItem.BackColor = default(Color);
                    eighthToolStripMenuItem.BackColor = default(Color);
                    sixteenthToolStripMenuItem.BackColor = default(Color);
                    break;
                case 2:
                    fullToolStripMenuItem.BackColor = default(Color);
                    halfToolStripMenuItem.BackColor = System.Drawing.Color.LightGray;
                    quarterToolStripMenuItem.BackColor = default(Color);
                    eighthToolStripMenuItem.BackColor = default(Color);
                    sixteenthToolStripMenuItem.BackColor = default(Color);
                    break;
                case 4:
                    fullToolStripMenuItem.BackColor = default(Color);
                    halfToolStripMenuItem.BackColor = default(Color);
                    quarterToolStripMenuItem.BackColor = System.Drawing.Color.LightGray;
                    eighthToolStripMenuItem.BackColor = default(Color);
                    sixteenthToolStripMenuItem.BackColor = default(Color);
                    break;
                case 8:
                    fullToolStripMenuItem.BackColor = default(Color);
                    halfToolStripMenuItem.BackColor = default(Color);
                    quarterToolStripMenuItem.BackColor = default(Color);
                    eighthToolStripMenuItem.BackColor = System.Drawing.Color.LightGray;
                    sixteenthToolStripMenuItem.BackColor = default(Color);
                    break;
                case 16:
                    fullToolStripMenuItem.BackColor = default(Color);
                    halfToolStripMenuItem.BackColor = default(Color);
                    quarterToolStripMenuItem.BackColor = default(Color);
                    eighthToolStripMenuItem.BackColor = default(Color);
                    sixteenthToolStripMenuItem.BackColor = System.Drawing.Color.LightGray;
                    break;
            }
            var coms = SerialPort.GetPortNames();
            port1 = new SerialPort();
            port1.ReadTimeout = 1500;

            ToolStripMenuItem[] items = new ToolStripMenuItem[coms.Length];

                for (int i = 0; i < items.Length; i++)
                {
                    items[i] = new ToolStripMenuItem();
                    items[i].Name = "dynamicItem" + i.ToString();
                    items[i].Tag = "specialData";
                    items[i].Text = coms[i];
                    items[i].Click += new EventHandler(PortMenuClickHandler);
                }

                portsToolStripMenuItem.DropDownItems.AddRange(items);

                void PortMenuClickHandler(object sender, EventArgs e)
                {
                    ToolStripMenuItem clickedItem = (ToolStripMenuItem)sender;

                try
                {
                    port1.PortName = clickedItem.ToString();
                    label6.Visible = true;
                    label6.Text = clickedItem.ToString();
                    for (int i = 0; i < items.Length; i++)
                    {
                        if (items[i] == clickedItem)
                        {
                            if (clickedItem.BackColor != System.Drawing.Color.LightGray)
                            {
                                clickedItem.BackColor = System.Drawing.Color.LightGray;
                            }
                        }
                        else { items[i].BackColor = default(Color); };
                    }
                }
                catch(Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                }
        }

        private void quitToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            Application.Exit();
            Environment.Exit(Environment.ExitCode);
        }


        private void toolStripMenuItem2_Click(object sender, EventArgs e)
        {
            ToolStripMenuItem clickedItem = (ToolStripMenuItem)sender;
            var baudrate = Convert.ToInt32(clickedItem.Text);
            port1.BaudRate = baudrate;
            clickedItem.BackColor= System.Drawing.Color.LightGray;
            toolStripMenuItem3.BackColor = default(Color);
            label1.Visible = true;
            label1.Text = clickedItem.Text;
        }

        private void toolStripMenuItem3_Click(object sender, EventArgs e)
        {
            ToolStripMenuItem clickedItem = (ToolStripMenuItem)sender;
            var baudrate = Convert.ToInt32(clickedItem.Text);
            port1.BaudRate = baudrate;
            clickedItem.BackColor = System.Drawing.Color.LightGray;
            toolStripMenuItem2.BackColor = default(Color);
            label1.Visible = true;
            label1.Text = clickedItem.Text;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (isPushed == false)
            {
                try
                {
                    label5.Text = "Connecting";
                    label1.Visible = true;
                    label1.Text = port1.BaudRate.ToString();
                    port1.Open();
                    try
                    {
                        
                        port1.Write(switchcode,0,1);
                        Thread.Sleep(1000);
                        int elapsed = 0;
                        while((connected==false)&& (elapsed<5000))
                        {
                            try
                            {
                                //Thread.Sleep(1000);
                                elapsed += 1000;
                                byte[] data = new byte[1024];
                                int bytesRead = port1.Read(data, 0, data.Length);
                                handshake = Encoding.ASCII.GetString(data, 0, bytesRead);
                                if (handshake=="hello winform\r\n")
                                {
                                    button1.Text = "Disconnect";
                                    isPushed = true;
                                    label5.Text = "Connected";
                                    connected = true;
                                    label12.Visible = true;
                                    label12.Text = "Ready";
                                    label16.Visible = true;
                                    label16.Text = "0";
                                }

                            }
                            catch (Exception ex){
                                MessageBox.Show(ex.Message);
                                label5.Text = "Disconnected";
                                break;
                            }
                        }


                    }
                    catch(Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                        label5.Text = "Disconnected";
                    }
                    if (!connected)
                    {
                        //port1.WriteLine("goodbye arduino");
                        port1.Close();
                        
                        label5.Text = "Disconnected";
                        isPushed = false;
                        button1.Text = "Connect";
                        label1.Visible = false;
                        MessageBox.Show("Could not connect to the microcontroller.");
                    }

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    label5.Text = "Disconnected";
                }
            }
            else
            {
                port1.Close();
                button1.Text = "Connect";
                isPushed = false;
                label5.Text = "Disconnected";
                label1.Visible = false;
                connected = false;
                label16.Visible = false;
                label12.Visible = false;
            }
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton1.Checked)
            {
                double area = Math.PI * Math.Pow(Properties.Settings.Default.defaultid1/2, 2);
                mlperstep = area*mmperstep * System.Math.Pow(10, -3)*Properties.Settings.Default.factor1;
                maxsteps = Properties.Settings.Default.defaultml1 / mlperstep;
            }
            else
            {
                mlperstep = -1;
                maxsteps = -1;
            }
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton2.Checked)
            {
                double area = Math.PI * Math.Pow(Properties.Settings.Default.defaultid2 / 2, 2);
                mlperstep = area * mmperstep * System.Math.Pow(10, -3) * Properties.Settings.Default.factor2;
                maxsteps = Properties.Settings.Default.defaultml2 / mlperstep;
            }
            else
            {
                mlperstep = -1;
                maxsteps = -1;
            }
        }

        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton3.Checked)
            {
                double area = Math.PI * Math.Pow(Properties.Settings.Default.defaultid3 / 2, 2);
                mlperstep = area * mmperstep * System.Math.Pow(10, -3) * Properties.Settings.Default.factor3;
                maxsteps = Properties.Settings.Default.defaultml3 / mlperstep;
            }
            else
            {
                mlperstep = -1;
                maxsteps = -1;
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            if (textBox1.TextLength > 0)
            {
                radioButton1.Enabled = false;
                radioButton2.Enabled = false;
                radioButton3.Enabled = false;
                string diameter = textBox1.Text;
                double diam = ConvertToDouble(diameter);
                double rad = diam / 2;
                double area = System.Math.PI*System.Math.Pow(rad,2); //mm^2
                mlperstep = mmperstep *area*System.Math.Pow(10,-3) * Properties.Settings.Default.factor4; 
            }
            else
            {
                radioButton1.Enabled = true;
                radioButton2.Enabled = true;
                radioButton3.Enabled = true;
                mlperstep = -1;
            }
        }
        private void textBox4_TextChanged(object sender, EventArgs e)
        {
            if (textBox4.TextLength > 0)
            {
                radioButton1.Enabled = false;
                radioButton2.Enabled = false;
                radioButton3.Enabled = false;
                string vol = textBox4.Text;
                double dvol = ConvertToDouble(vol);
                if ((mlperstep != -1)&&(textBox1.TextLength>0))
                {
                    maxsteps = dvol / mlperstep;
                }
                else
                {
                    radioButton1.Enabled = true;
                    radioButton2.Enabled = true;
                    radioButton3.Enabled = true;
                    maxsteps = -1;
                    MessageBox.Show("Please specify custom syringe diameter first.");
                    textBox4.Text = "";
                }
            }
            else
            {
                radioButton1.Enabled = true;
                radioButton2.Enabled = true;
                radioButton3.Enabled = true;
                maxsteps = -1;
            }
        }


        private void button4_Click(object sender, EventArgs e)
        {
            if (connected)
            {
                port1.Write(switchcode,1,1);
                label12.Text = "Homing";
                bool homed = false;
                while (!homed)
                {
                    try
                    {
                        if (port1.BytesToRead > 0)
                        {
                            Thread.Sleep(2000);
                            byte[] data = new byte[1024];
                            int bytesRead = port1.Read(data, 0, data.Length);
                            string pos = Encoding.ASCII.GetString(data, 0, bytesRead);
                            if (pos == "yup hom")
                            {
                                homed = true;
                                ishomed = true;
                                label12.Text = "Ready";
                                position = 0;
                                crntvlm = 0;
                                label16.Text = crntvlm.ToString();
                            }
                        }
                    }
                    catch (Exception ex){
                        MessageBox.Show(ex.Message);
                        break;
                    }
                }
            }
        }

        delegate void SetTextCallback(string text);

        private void SetText(string text)
        { 
            if (label16.InvokeRequired)
            {
                SetTextCallback d = new SetTextCallback(SetText);
                this.Invoke(d, new object[] { text });
            }
            else
            {
                label16.Text = text;
            }
        }

        private void button3_Click(object sender, EventArgs e)//inject
        {
            if (connected)
            {
                if (!isPushedinj)
                {
                    if ((textBox2.TextLength > 0) && (textBox3.TextLength > 0))
                    {
                        vlm = ConvertToDouble(textBox2.Text);
                        rate = ConvertToDouble(textBox3.Text);
                        try
                        {
                            if ((mlperstep != -1) && (maxsteps != -1))
                            {
                                if (ishomed)
                                {
                                    outputposition = vlm / mlperstep;
                                    outputrate = (rate / 60) / mlperstep;

                                    if (outputrate <= 100000)
                                    {
                                        if (position - outputposition >= 0)
                                        {
                                            double movetoposition = position - outputposition;
                                            crntvlm = movetoposition * mlperstep;
                                            crntvlm = Math.Round(crntvlm, 4, MidpointRounding.AwayFromZero);

                                            int actualposition = (int)Math.Round(outputposition, 0, MidpointRounding.AwayFromZero);
                                            //position = position - actualposition;
                                            string actpsn = actualposition.ToString();

                                            int actualrate = (int)Math.Round(outputrate, 0, MidpointRounding.AwayFromZero);
                                            string actrt = actualrate.ToString();
                                            if (actpsn.Length == 6)
                                            {
                                                contpsn = "x0" + actpsn;
                                            }
                                            else if (actpsn.Length == 5)
                                            {
                                                contpsn = "x00" + actpsn;
                                            }
                                            else if (actpsn.Length == 4)
                                            {
                                                contpsn = "x000" + actpsn;
                                            }
                                            else if (actpsn.Length == 3)
                                            {
                                                contpsn = "x0000" + actpsn;
                                            }
                                            else if (actpsn.Length == 2)
                                            {
                                                contpsn = "x00000" + actpsn;
                                            }
                                            else if (actpsn.Length == 1)
                                            {
                                                contpsn = "x000000" + actpsn;
                                            }
                                            else
                                            {
                                                MessageBox.Show("Volume out of range!");
                                            }
                                            if (actrt.Length == 5)
                                            {
                                                contrt = "v0" + actrt;
                                            }
                                            else if (actrt.Length == 4)
                                            {
                                                contrt = "v00" + actrt;
                                            }
                                            else if (actrt.Length == 3)
                                            {
                                                contrt = "v000" + actrt;
                                            }
                                            else if (actrt.Length == 2)
                                            {
                                                contrt = "v0000" + actrt;
                                            }
                                            else if (actrt.Length == 1)
                                            {
                                                contrt = "v00000" + actrt;
                                            }
                                            else
                                            {
                                                MessageBox.Show("Rate out of range!");
                                            }
                                            if ((contpsn.Length == 8) && (contrt.Length == 7))
                                            {
                                                sendit = contpsn + contrt + "x2";
                                                port1.Write(switchcode, 2, 1);
                                                byte[] sendbytes = Encoding.GetEncoding("ASCII").GetBytes(sendit);
                                                port1.Write(sendbytes, 0, sendbytes.Length);
                                                watch.Reset();
                                                watch.Start();
                                                int interval = actualposition / actualrate;
                                                timer.Interval = interval * 1000; // 
                                                timer.Tick += timer_Tick;
                                                timer.Start();
                                                
                                                isPushedinj = true;
                                                injected = true;
                                                label12.Text = "Injecting";
                                                button1.Enabled = false;
                                                button3.Text = "Cancel";
                                                button2.Enabled = false;
                                                button4.Enabled = false;
                                            }
                                        }
                                        else
                                        {
                                            MessageBox.Show("Injected volume would be bigger than current volume! :(");
                                        }
                                    }
                                    else
                                    {
                                        MessageBox.Show("Rate too high!");
                                    }
                                }
                                else { MessageBox.Show("Pump position not homed!"); }
                            }
                            else { MessageBox.Show("No Syringe Chosen!"); }

                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message);
                        }

                    }
                    else
                    {
                        MessageBox.Show("Invalid Input");
                    }
                }
                else
                {
                    port1.Write(switchcode, 3, 1);
                    button3.Text = "Inject";
                    button1.Enabled = true;
                    button2.Enabled = true;
                    
                    button4.Enabled = true;
                    injected = false;
                    long min = watch.ElapsedMilliseconds;
                    watch.Stop();
                    timer.Stop();

                    double ratepermilisec = rate / (1000 * 60);
                    double vlm = ratepermilisec * min;
                    double stepstaken = vlm / mlperstep;
                    position = position - stepstaken;
                    double newvlm = position * mlperstep;
                    crntvlm = Math.Round(newvlm, 4, MidpointRounding.AwayFromZero);
                    
                    label16.Text = crntvlm.ToString();
                    label12.Text = "Ready";
                    isPushedinj = false;
                }
            }
        }
        private double ConvertToDouble(string s)
        {
            char systemSeparator = Thread.CurrentThread.CurrentCulture.NumberFormat.CurrencyDecimalSeparator[0];
            double result = 0;
            try
            {
                if (s != null)
                    if (!s.Contains(","))
                        result = double.Parse(s, CultureInfo.InvariantCulture);
                    else
                        result = Convert.ToDouble(s.Replace(".", systemSeparator.ToString()).Replace(",", systemSeparator.ToString()));
            }
            catch (Exception e)
            {
                try
                {
                    result = Convert.ToDouble(s);
                }
                catch
                {
                    try
                    {
                        result = Convert.ToDouble(s.Replace(",", ";").Replace(".", ",").Replace(";", "."));
                    }
                    catch
                    {
                        MessageBox.Show("Wrong format, input must be numeric!");
                        textBox2.Text = "";
                        textBox3.Text = "";
                    }
                }
            }
            return result;
        }

        private static DialogResult ShowInputDialog(ref string input)
        {
            System.Drawing.Size size = new System.Drawing.Size(200, 70);
            Form inputBox = new Form();

            inputBox.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            inputBox.ClientSize = size;
            inputBox.Text = "Input";

            System.Windows.Forms.TextBox textBox = new TextBox();
            textBox.Size = new System.Drawing.Size(size.Width - 10, 23);
            textBox.Location = new System.Drawing.Point(5, 5);
            textBox.Text = input;
            inputBox.Controls.Add(textBox);

            Button okButton = new Button();
            okButton.DialogResult = System.Windows.Forms.DialogResult.OK;
            okButton.Name = "okButton";
            okButton.Size = new System.Drawing.Size(75, 23);
            okButton.Text = "&OK";
            okButton.Location = new System.Drawing.Point(size.Width - 80 - 80, 39);
            inputBox.Controls.Add(okButton);

            Button cancelButton = new Button();
            cancelButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            cancelButton.Name = "cancelButton";
            cancelButton.Size = new System.Drawing.Size(75, 23);
            cancelButton.Text = "&Cancel";
            cancelButton.Location = new System.Drawing.Point(size.Width - 80, 39);
            inputBox.Controls.Add(cancelButton);

            inputBox.AcceptButton = okButton;
            inputBox.CancelButton = cancelButton;

            DialogResult result = inputBox.ShowDialog();
            input = textBox.Text;
            return result;
        }
        private static DialogResult ShowCalibrationDialog(params string[] paramArray)
        {
            System.Drawing.Size size = new System.Drawing.Size(270, 150);
            Form inputBox = new Form();

            inputBox.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            inputBox.ClientSize = size;
            inputBox.Text = "Calibrate";

            System.Windows.Forms.Label label1=new Label();
            label1.Location = new System.Drawing.Point(1, 5);
            label1.Text = "Average measured volume (mL)";
            label1.Size = new System.Drawing.Size(170,13);
            inputBox.Controls.Add(label1);

            System.Windows.Forms.Label label2 = new Label();
            label2.Location = new System.Drawing.Point(1, 50);
            label2.Text = "Theoretical delivered volume (mL)";
            label2.Size = new System.Drawing.Size(170, 13);
            inputBox.Controls.Add(label2);

            System.Windows.Forms.TextBox textBox1 = new TextBox();
            textBox1.Size = new System.Drawing.Size(80, 23);
            textBox1.Location = new System.Drawing.Point(180, 5);
            inputBox.Controls.Add(textBox1);

            System.Windows.Forms.TextBox textBox2 = new TextBox();
            textBox2.Size = new System.Drawing.Size(80, 23);
            textBox2.Location = new System.Drawing.Point(180, 50);
            inputBox.Controls.Add(textBox2);

            Button okButton = new Button();
            okButton.DialogResult = System.Windows.Forms.DialogResult.OK;
            okButton.Name = "okButton";
            okButton.Size = new System.Drawing.Size(75, 23);
            okButton.Text = "&OK";
            okButton.Location = new System.Drawing.Point(size.Width - 82 - 82, 80);
            inputBox.Controls.Add(okButton);

            Button cancelButton = new Button();
            cancelButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            cancelButton.Name = "cancelButton";
            cancelButton.Size = new System.Drawing.Size(75, 23);
            cancelButton.Text = "&Cancel";
            cancelButton.Location = new System.Drawing.Point(size.Width - 84, 80);
            inputBox.Controls.Add(cancelButton);

            inputBox.AcceptButton = okButton;
            inputBox.CancelButton = cancelButton;

            DialogResult result = inputBox.ShowDialog();
            paramArray[0] = textBox1.Text;
            paramArray[1] = textBox2.Text;
            
            return result;
        }
        private void button2_Click(object sender, EventArgs e)//aspirate
        {
            if (connected)
            {
                if (!isPushedasp)
                {
                    if ((textBox2.TextLength > 0) && (textBox3.TextLength > 0))
                    {
                        vlm = ConvertToDouble(textBox2.Text);
                        rate = ConvertToDouble(textBox3.Text);
                        try
                        {

                            if ((mlperstep != -1) && (maxsteps != -1))
                            {
                                if (ishomed)
                                {
                                    outputrate = (rate / 60) / mlperstep;
                                    outputposition = vlm / mlperstep;
                                    //position = 0;
                                    if (outputrate <= 100000)
                                    {
                                        if (position + outputposition <= maxsteps)
                                        {
                                            
                                            double movetoposition = position + outputposition;
                                            crntvlm = movetoposition * mlperstep;
                                            crntvlm = Math.Round(crntvlm, 4, MidpointRounding.AwayFromZero);
                                            //crntvlm = (int)Math.Round(dbcrntvlm, 0, MidpointRounding.AwayFromZero);
                                            int actualposition = (int)Math.Round(outputposition, 0, MidpointRounding.AwayFromZero);
                                            //position = position + actualposition;
                                            string actpsn = actualposition.ToString();

                                            int actualrate = (int)Math.Round(outputrate, 0, MidpointRounding.AwayFromZero);
                                            string actrt = actualrate.ToString();
                                            if (actpsn.Length == 6)
                                            {
                                                contpsn = "x0" + actpsn;
                                            }
                                            else if (actpsn.Length == 5)
                                            {
                                                contpsn = "x00" + actpsn;
                                            }
                                            else if (actpsn.Length == 4)
                                            {
                                                contpsn = "x000" + actpsn;
                                            }
                                            else if (actpsn.Length == 3)
                                            {
                                                contpsn = "x0000" + actpsn;
                                            }
                                            else if (actpsn.Length == 2)
                                            {
                                                contpsn = "x00000" + actpsn;
                                            }
                                            else if (actpsn.Length == 1)
                                            {
                                                contpsn = "x000000" + actpsn;
                                            }
                                            else
                                            {
                                                MessageBox.Show("Volume out of range!");
                                            }
                                            if (actrt.Length == 5)
                                            {
                                                contrt = "v0" + actrt;
                                            }
                                            else if (actrt.Length == 4)
                                            {
                                                contrt = "v00" + actrt;
                                            }
                                            else if (actrt.Length == 3)
                                            {
                                                contrt = "v000" + actrt;
                                            }
                                            else if (actrt.Length == 2)
                                            {
                                                contrt = "v0000" + actrt;
                                            }
                                            else if (actrt.Length == 1)
                                            {
                                                contrt = "v00000" + actrt;
                                            }
                                            else
                                            {
                                                MessageBox.Show("Rate out of range!");
                                            }
                                            if ((contpsn.Length == 8) && (contrt.Length == 7))
                                            {
                                                sendit = contpsn + contrt + "x1";
                                                port1.Write(switchcode, 2, 1);
                                                
                                                byte[] sendbytes = Encoding.GetEncoding("ASCII").GetBytes(sendit);
                                                port1.Write(sendbytes, 0, sendbytes.Length);
                                                
                                                int interval = actualposition / actualrate;
                                                timer.Interval = interval * 1000; // here time in milliseconds
                                                timer.Tick += timer_Tick;
                                                timer.Start();
                                                watch.Reset();
                                                watch.Start();
                                                isPushedasp = true;
                                                aspirated = true;
                                                label12.Text = "Aspirating";
                                                button1.Enabled = false;
                                                button2.Text = "Cancel";
                                                button3.Enabled = false;
                                                button4.Enabled = false;
                                            }
                                        }
                                        else
                                        {
                                            MessageBox.Show("Aspirated volume would be bigger than syringe volume! :(");
                                        }
                                    }
                                    else
                                    {
                                        MessageBox.Show("Rate too high!");
                                    }
                                }
                                else { MessageBox.Show("Pump position not homed!"); }
                            }
                            else { MessageBox.Show("No Syringe Chosen!"); }

                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message);
                        }

                    }
                    else
                    {
                        MessageBox.Show("Invalid Input");
                    }
                }
                else
                {

                    port1.Write(switchcode, 3, 1);
                    button2.Text = "Aspirate";
                    aspirated = false;
                    button1.Enabled = true;
                    
                    button3.Enabled = true;
                    button4.Enabled = true;
                    long min = watch.ElapsedMilliseconds;
                    double ratepermilisec = rate /( 1000 * 60);
                    //MessageBox.Show(min.ToString());
                    watch.Stop();
                    timer.Stop();
                    
                    double vlm = ratepermilisec * min;
                    double stepstaken = vlm / mlperstep;
                    position = position + stepstaken;
                    double newvlm = position * mlperstep;
                    crntvlm = Math.Round(newvlm, 4, MidpointRounding.AwayFromZero);
                    label16.Text = crntvlm.ToString();
                    label12.Text = "Ready";
                    isPushedasp = false;
                }
            }
        }

        private void mL(object sender, EventArgs e)
        {

        }

        private void defaultToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (connected)
            {
                ShowInputDialog(ref usrinput);
                if (usrinput.Length > 0)
                {
                    double newaccel = ConvertToDouble(usrinput);
                    Properties.Settings.Default.accel = newaccel;
                    Properties.Settings.Default.Save();
                    defaultStepsecToolStripMenuItem.Text =usrinput;
                    if (newaccel > 0)
                    {
                        int actualaccel = (int)Math.Round(newaccel, 0, MidpointRounding.AwayFromZero) *Properties.Settings.Default.microstp;
                        if (actualaccel < 20000)
                        {
                            port1.Write(switchcode, 4, 1);
                            Thread.Sleep(1000);
                            string accelstr = actualaccel.ToString();
                            if (accelstr.Length == 5) { accelsend = "g" + accelstr; }
                            else if (accelstr.Length == 4) { accelsend = "g0" + accelstr; }
                            else if (accelstr.Length == 3) { accelsend = "g00" + accelstr; }
                            else if (accelstr.Length == 2) { accelsend = "g000" + accelstr; }
                            else if (accelstr.Length == 1) { accelsend = "g0000" + accelstr; }
                            else { MessageBox.Show("Acceleration too small!"); }
                            if (accelsend.Length == 6)
                            {
                                port1.Write(accelsend);
                            }
                            bool accelchanged = false;
                            int aceltime = 0;
                            while (!accelchanged)
                            {
                                
                                if (aceltime < 5000)
                                {
                                    try
                                    {

                                        if (port1.BytesToRead > 0)
                                        {
                                            //Thread.Sleep(1000);
                                            byte[] data = new byte[1024];
                                            int bytesRead = port1.Read(data, 0, data.Length);
                                            string pos = Encoding.ASCII.GetString(data, 0, bytesRead);
                                            if (pos == "yup dun")
                                            {
                                                accelchanged = true;
                                                MessageBox.Show("Acceleration updated!");
                                                customaccel = true;
                                            }
                                        }
                                        aceltime = aceltime + 1000;
                                        Thread.Sleep(1000);
                                    }
                                    catch (Exception ex)
                                    {
                                        MessageBox.Show(ex.Message);
                                        break;
                                    }
                                }
                                else {
                                    MessageBox.Show("Could not update acceleration!");
                                    break;
                                }
                            }
                        }
                        else { MessageBox.Show("Slow down professor!"); }
                    }
                }
            }

            else { MessageBox.Show("Please connect to the pump before changing anything."); }
        }
        void timer_Tick(object sender, System.EventArgs e)
        {
            button1.Enabled = true;
            button2.Enabled = true;
            button3.Enabled = true;
            button4.Enabled = true;
            isPushedasp = false;
            isPushedinj = false;
            button2.Text = "Aspirate";
            button3.Text = "Inject";
            label16.Text = crntvlm.ToString();
            label12.Text = "Ready";
            timer.Stop();
            watch.Stop();
            if (aspirated) {
                position = position + outputposition;
                    }else if (injected)
            {
                position = position - outputposition;
            }
            aspirated = false;
            injected = false;
        }
        

        private void rotationPerStepdegreesToolStripMenuItem_Click(object sender, EventArgs e)
        {
           ShowInputDialog(ref usrinput);
            if (usrinput.Length > 0)
            {
                double newdegree = ConvertToDouble(usrinput);
                Properties.Settings.Default.degree = newdegree;
                Properties.Settings.Default.Save();
                rotationPerStepdegreesToolStripMenuItem.Text = "Stepper degree: "+ newdegree.ToString();
            }
        }

        private void pitchToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ShowInputDialog(ref usrinput);
            if (usrinput.Length > 0)
            {
                double newpitch = ConvertToDouble(usrinput);
                Properties.Settings.Default.pitch = newpitch;
                Properties.Settings.Default.Save();
                pitchToolStripMenuItem.Text = "Ball screw pitch= "+newpitch.ToString() + " mm";
                mmperstep = Properties.Settings.Default.pitch / ((360 / Properties.Settings.Default.degree) * Properties.Settings.Default.microstp);
            }
        }

        private void innerDiameterToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ShowInputDialog(ref usrinput);
            if (usrinput.Length > 0)
            {
                double newid = ConvertToDouble(usrinput);
                Properties.Settings.Default.defaultid1 = newid;
                Properties.Settings.Default.Save();
               innerDiameterToolStripMenuItem.Text = "Inner diameter= " + newid.ToString() + " mm";
            }
        }

        private void innerDiameterToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            ShowInputDialog(ref usrinput);
            if (usrinput.Length > 0)
            {
                double newid = ConvertToDouble(usrinput);
                Properties.Settings.Default.defaultid2 = newid;
                Properties.Settings.Default.Save();
                innerDiameterToolStripMenuItem1.Text = "Inner diameter= " + newid.ToString() + " mm";
            }
        }

        private void innerDiameterToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            ShowInputDialog(ref usrinput);
            if (usrinput.Length > 0)
            {
                double newid = ConvertToDouble(usrinput);
                Properties.Settings.Default.defaultid3 = newid;
                Properties.Settings.Default.Save();
                innerDiameterToolStripMenuItem2.Text = "Inner diameter= " + newid.ToString() + " mm";
            }
        }

        private void fullToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.microstp = 1;
            fullToolStripMenuItem.BackColor = System.Drawing.Color.LightGray;
            halfToolStripMenuItem.BackColor = default(Color);
            quarterToolStripMenuItem.BackColor= default(Color);
            eighthToolStripMenuItem.BackColor= default(Color);
            sixteenthToolStripMenuItem.BackColor= default(Color);
        }

        private void halfToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.microstp = 2;
            fullToolStripMenuItem.BackColor = default(Color);
            halfToolStripMenuItem.BackColor = System.Drawing.Color.LightGray;
            quarterToolStripMenuItem.BackColor = default(Color);
            eighthToolStripMenuItem.BackColor = default(Color);
            sixteenthToolStripMenuItem.BackColor = default(Color);
        }

        private void quarterToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.microstp = 4;
            fullToolStripMenuItem.BackColor = default(Color);
            halfToolStripMenuItem.BackColor = default(Color);
            quarterToolStripMenuItem.BackColor = System.Drawing.Color.LightGray;
            eighthToolStripMenuItem.BackColor = default(Color);
            sixteenthToolStripMenuItem.BackColor = default(Color);
        }

        private void eighthToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.microstp = 8;
            fullToolStripMenuItem.BackColor = default(Color);
            halfToolStripMenuItem.BackColor = default(Color);
            quarterToolStripMenuItem.BackColor = default(Color);
            eighthToolStripMenuItem.BackColor = System.Drawing.Color.LightGray;
            sixteenthToolStripMenuItem.BackColor = default(Color);
        }

        private void sixteenthToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.microstp = 8;
            fullToolStripMenuItem.BackColor = default(Color);
            halfToolStripMenuItem.BackColor = default(Color);
            quarterToolStripMenuItem.BackColor = default(Color);
            eighthToolStripMenuItem.BackColor = default(Color);
            sixteenthToolStripMenuItem.BackColor = System.Drawing.Color.LightGray;
        }

        private void changeVolumeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ShowInputDialog(ref usrinput);
            if (usrinput.Length > 0)
            {
                double newid = ConvertToDouble(usrinput);
                Properties.Settings.Default.defaultml1 = newid;
                Properties.Settings.Default.Save();
                mLToolStripMenuItem.Text = Properties.Settings.Default.defaultml1.ToString() + " mL";
                mLToolStripMenuItem3.Text = Properties.Settings.Default.defaultml1.ToString() + " mL";
                radioButton1.Text= Properties.Settings.Default.defaultml1.ToString() + " mL";
            }
        }

        private void changeVolumeToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            ShowInputDialog(ref usrinput);
            if (usrinput.Length > 0)
            {
                double newid = ConvertToDouble(usrinput);
                Properties.Settings.Default.defaultml2 = newid;
                Properties.Settings.Default.Save();
                mLToolStripMenuItem1.Text = Properties.Settings.Default.defaultml2.ToString() + " mL";
                mLToolStripMenuItem4.Text = Properties.Settings.Default.defaultml1.ToString() + " mL";
                radioButton2.Text = Properties.Settings.Default.defaultml1.ToString() + " mL";
            }
        }

        private void changeVolumeToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            ShowInputDialog(ref usrinput);
            if (usrinput.Length > 0)
            {
                double newid = ConvertToDouble(usrinput);
                Properties.Settings.Default.defaultml3 = newid;
                Properties.Settings.Default.Save();
                mLToolStripMenuItem2.Text = Properties.Settings.Default.defaultml3.ToString() + " mL";
                mLToolStripMenuItem5.Text = Properties.Settings.Default.defaultml1.ToString() + " mL";
                radioButton3.Text = Properties.Settings.Default.defaultml1.ToString() + " mL";
            }
        }

        private void mLToolStripMenuItem3_Click(object sender, EventArgs e)
        {
            string actvl=null;
            string theovl=null;
            string[] inputs = new string[] { actvl, theovl};
            ShowCalibrationDialog(inputs);
            if((inputs[0].Length > 0) && (inputs[1].Length > 0))
            {
                double act = ConvertToDouble(inputs[0]);
                double theo = ConvertToDouble(inputs[1]);
                Properties.Settings.Default.factor1 = act / theo;
                Properties.Settings.Default.Save();
               
                MessageBox.Show("Calibration saved!");
            }
        }

        private void mLToolStripMenuItem4_Click(object sender, EventArgs e)
        {
            string actvl = null;
            string theovl = null;
            string[] inputs = new string[] { actvl, theovl };
            ShowCalibrationDialog(inputs);
            if ((inputs[0].Length > 0) && (inputs[1].Length > 0))
            {
                double act = ConvertToDouble(inputs[0]);
                double theo = ConvertToDouble(inputs[1]);
                Properties.Settings.Default.factor2 = act / theo;
                Properties.Settings.Default.Save();
               
                MessageBox.Show("Calibration saved!");
            }
        }

        private void mLToolStripMenuItem5_Click(object sender, EventArgs e)
        {
            string actvl = null;
            string theovl = null;
            string[] inputs = new string[] { actvl, theovl };
            ShowCalibrationDialog(inputs);
            if ((inputs[0].Length > 0) && (inputs[1].Length > 0))
            {
                double act = ConvertToDouble(inputs[0]);
                double theo = ConvertToDouble(inputs[1]);
                Properties.Settings.Default.factor3 = act / theo;
                Properties.Settings.Default.Save();
 
                MessageBox.Show("Calibration saved!");
            }
        }

        private void calibrateToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string actvl = null;
            string theovl = null;
            string[] inputs = new string[] { actvl, theovl };
            ShowCalibrationDialog(inputs);
            if ((inputs[0].Length > 0) && (inputs[1].Length > 0))
            {
                double act = ConvertToDouble(inputs[0]);
                double theo = ConvertToDouble(inputs[1]);
                Properties.Settings.Default.factor4 = act / theo;
                Properties.Settings.Default.Save();

                MessageBox.Show("Calibration saved!");
            }
        }

        private void refreshToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var coms = SerialPort.GetPortNames();
            ToolStripMenuItem[] items = new ToolStripMenuItem[coms.Length];

            for (int i = 0; i < items.Length; i++)
            {
                items[i] = new ToolStripMenuItem();
                items[i].Name = "dynamicItem" + i.ToString();
                items[i].Tag = "specialData";
                items[i].Text = coms[i];
                items[i].Click += new EventHandler(PortMenuClickHandler);
            }
            for (int j=0; j<items.Length; j++) {
                for (int v=1; v<portsToolStripMenuItem.DropDownItems.Count;v++) {
                    if (items[j].Text != portsToolStripMenuItem.DropDownItems[v].ToString())
                    {
                        portsToolStripMenuItem.DropDownItems.Add(items[j]);
                    }
                }
            }
            //portsToolStripMenuItem.DropDownItems.AddRange(items);

            void PortMenuClickHandler(object sender2, EventArgs e2)
            {
                ToolStripMenuItem clickedItem = (ToolStripMenuItem)sender2;

                try
                {
                    port1.PortName = clickedItem.ToString();
                    label6.Visible = true;
                    label6.Text = clickedItem.ToString();
                    for (int i = 0; i < items.Length; i++)
                    {
                        if (items[i] == clickedItem)
                        {
                            if (clickedItem.BackColor != System.Drawing.Color.LightGray)
                            {
                                clickedItem.BackColor = System.Drawing.Color.LightGray;
                            }
                        }
                        else { items[i].BackColor = default(Color); };
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }
    }
}
