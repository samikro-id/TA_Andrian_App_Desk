using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using MQTTnet;
using MQTTnet.Client;
using MQTTnet.Client.Options;
using System.IO.Ports;

namespace TA_Andrian_App_Desk
{
    public partial class fmMain : Form
    {
        private static IMqttClient client;
        private static IMqttClientOptions options;

        private const string mqtt_url = "broker.emqx.io";
        private const int mqtt_port = 1883;
        private const string mqtt_username = "";
        private const string mqtt_password = "";
        private const string mqtt_id = "andrianAppDesk";

        private const string start_form = "START";
        private const string stop_form = "STOP";

        private const string cmd_data = "GET|DATA";
        private const string cmd_source_pln = "SET|SOURCE|PLN";
        private const string cmd_source_pltph = "SET|SOURCE|PLTPH";
        private const string cmd_protec_on = "SET|PROTEC|ON";
        private const string cmd_protec_off = "SET|PROTEC|OFF";

        public fmMain()
        {
            InitializeComponent();

            //webPowerOut.ObjectForScripting = this;
        }

        /* MQTT Section */
        public static void publishMqtt(string payload)
        {
            try
            {
                if (client.IsConnected || client.IsConnected != null)
                {
                    var message = new MqttApplicationMessageBuilder()
                    .WithTopic("samikro/cmd/project/3")
                    .WithPayload(payload)
                    .WithExactlyOnceQoS()
                    .Build();

                    client.PublishAsync(message);
                }
            }
            catch(Exception exc)
            {
                Console.WriteLine(exc.Message);
            }
        }

        public static Task SubscribeAsync(string _topic, int qos = 1) =>
              client.SubscribeAsync(new MqttTopicFilterBuilder()
                .WithTopic(_topic)
                .WithQualityOfServiceLevel((MQTTnet.Protocol.MqttQualityOfServiceLevel)qos)
                .Build());

        public void connectedHandlerMqtt()
        {
            Console.WriteLine("MQTT is Connected");

            SubscribeAsync("samikro/data/project/3");

            btStartStop.Text = stop_form;
            plSetting.Enabled = false;
            plButton.Enabled = true;
            tmUpdateData.Enabled = true;
        }

        public void disconnectedHandlerMqtt()
        {
            Console.WriteLine("Mqtt Disconnect");

            btStartStop.Text = start_form;
            plButton.Enabled = false;
            plSetting.Enabled = true;
            tmUpdateData.Enabled = false;
        }

        public bool connectingMqtt()
        {
            Console.WriteLine("Connecting to MQTT");
            try
            {
                var factory = new MqttFactory();
                client = factory.CreateMqttClient();

                options = new MqttClientOptionsBuilder()
                    .WithClientId(mqtt_id + Guid.NewGuid().ToString())
                    .WithTcpServer(mqtt_url, mqtt_port)
                    .WithCredentials(mqtt_username, mqtt_password)
                    .WithCleanSession()
                    .Build();

                client.UseConnectedHandler(ex => {
                    this.Invoke((MethodInvoker)delegate ()
                    {
                        connectedHandlerMqtt();
                    });
                });

                client.UseDisconnectedHandler(ex => {
                    this.Invoke((MethodInvoker)delegate ()
                    {
                        disconnectedHandlerMqtt();
                    });
                });
                client.UseApplicationMessageReceivedHandler(ex => {
                    try
                    {
                        string payload = Encoding.UTF8.GetString(ex.ApplicationMessage.Payload);

                        this.Invoke((MethodInvoker)delegate ()
                        {
                            Console.WriteLine("Mqtt Message");
                            Console.WriteLine(payload);

                            parse_data(payload);
                        });
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine("Mqtt Message is Bad!!");
                        Console.WriteLine(e);
                    }
                });

                client.ConnectAsync(options);
                Console.WriteLine("MQTT Binding...");

                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine("MQTT Failed");
                Console.WriteLine(ex);

                return false;
            }
        }

        public void disconnectMqtt()
        {
            if (client.IsConnected)
            {
                Console.WriteLine("Disconnect Mqtt");
                client.Dispose();
                client.DisconnectAsync();
            }
        }
        /* End Of Mqtt Section */

        /* Serial Section */
        private void serial_open()
        {
            try
            {
                spSerial.PortName = cbSerial.SelectedItem.ToString();
                spSerial.BaudRate = 9600;
                spSerial.DataBits = 8;
                spSerial.StopBits = StopBits.One;
                spSerial.Parity = Parity.None;
                spSerial.ReadTimeout = 500;
                spSerial.WriteTimeout = 500;

                spSerial.Open();
            }
            catch(Exception exc)
            {
                Console.WriteLine(exc.Message);
                lbNotification.Text = "Failed open serial";
            }
        }

        private void serial_close()
        {
            spSerial.Dispose();
            spSerial.Close();
        }

        private void serial_send(string data)
        {
            if (serial_check())
            {
                spSerial.ReadExisting();

                Console.WriteLine("ss : " + data);

                spSerial.WriteLine(data);
            }
        }

        private void serial_dataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            try
            {
                String data = spSerial.ReadLine();
                data = data.Replace("\n", "").Replace("\r", "");

                Console.WriteLine("Serial Message");
                Console.WriteLine(data);

                this.BeginInvoke(new Action(() =>
                {
                    this.parse_data(data);
                }));
            }
            catch (Exception exc)
            {
                Console.WriteLine(exc.Message);
            }
        }

        private bool serial_check()
        {
            if (spSerial.IsOpen)
            {
                return true;
            }
            else
            {
                lbNotification.Text = "Serial port is close";
                return false;
            }
        }

        private void cbSerial_refresh()
        {
            cbSerial.Items.Clear();
            cbSerial.Text = "";
            cbSerial.Items.AddRange(SerialPort.GetPortNames());
        }
        /* End Of Serial Section */

        private void parse_data(String data)
        {
            try
            {
                string[] parsing = data.Split('|');
                Console.WriteLine(parsing[0]);

                if (parsing[0] == "DATA")
                {
                    lbTemperatureVal.Text = parsing[1];

                    lbSource.Text = (parsing[14] == "1") ? "PLN" : "PLTPH";

                    if(parsing[14] == "1")
                    {
                        lbPlnVoltageVal.Text = parsing[2];
                        lbPlnCurrentVal.Text = parsing[3];
                        lbPlnPowerVal.Text = parsing[4];
                        lbPlnEnergyVal.Text = parsing[7];

                        lbPltphVoltage.Text = "0";
                        lbPltphCurrent.Text = "0";
                        lbPltphPower.Text = "0";
                        lbPltphEnergy.Text = "0";
                    }
                    else
                    {
                        lbPlnVoltageVal.Text = "0";
                        lbPlnCurrentVal.Text = "0";
                        lbPlnPowerVal.Text = "0";
                        lbPlnEnergyVal.Text = "0";

                        lbPltphVoltage.Text = parsing[2];
                        lbPltphCurrent.Text = parsing[3];
                        lbPltphPower.Text = parsing[4];
                        lbPltphEnergy.Text = parsing[7];
                    }

                    lbOutputVoltage.Text = parsing[8];
                    lbOutputCurrent.Text = parsing[9];
                    lbOutputPower.Text = parsing[10];
                    lbOutputEnergy.Text = parsing[13];

                    
                    lbProtection.Text = (parsing[15] == "1") ? "ON" : "OFF";
                    lbFanVal.Text = (parsing[16] == "1") ? "ON" : "OFF";
                    lbOverTemp.Text = (parsing[17] == "1") ? "ON" : "OFF";
                    lbLowVolt.Text = (parsing[18] == "1") ? "ON" : "OFF";


                }
                else if (parsing[0] == "SOURCE")
                {
                    lbSource.Text = parsing[1];
                }
                else if(parsing[0] == "PROTEC")
                {
                    lbProtection.Text = parsing[1];
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("parse failed");
                Console.WriteLine(e);
            }
        }

        private void request_data() {
            send_data(cmd_data);
        }

        private void send_data(string data)
        {
            if (rbInternet.Checked)
            {
                publishMqtt(data);
            }
            else if (rbSerial.Checked)
            {
                serial_send(data);
            }
        }

        private void tmUpdateData_Tick(object sender, EventArgs e)
        {
            request_data();
        }

        private void btExit_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void rbSerial_CheckedChanged(object sender, EventArgs e)
        {
            cbSerial.Enabled = rbSerial.Checked;
        }

        private void btStartStop_Click(object sender, EventArgs e)
        {
            if(btStartStop.Text == start_form)
            {
                if (rbInternet.Checked)
                {
                    connectingMqtt();
                }
                else if (rbSerial.Checked)
                {
                    if(cbSerial.Text != "")
                    {
                        serial_open();

                        if (spSerial.IsOpen)
                        {
                            cbSerial.Enabled = false;

                            btStartStop.Text = stop_form;
                        }
                    }
                    else
                    {
                        lbNotification.Text = "Select serial port !!!";
                    }
                }

                if(btStartStop.Text == stop_form)
                {
                    plSetting.Enabled = false;
                    plButton.Enabled = true;
                    tmUpdateData.Enabled = true;
                }
            }
            else
            {
                if (rbInternet.Checked)
                {
                    disconnectMqtt();

                    btStartStop.Text = start_form;
                }
                else if (rbSerial.Checked)
                {
                    serial_close();

                    if (!spSerial.IsOpen)
                    {
                        cbSerial.Enabled = true;

                        btStartStop.Text = start_form;
                    }
                }

                if (btStartStop.Text == start_form)
                {
                    plSetting.Enabled = true;
                    plButton.Enabled = false;
                    tmUpdateData.Enabled = false;
                }
            }
        }

        private void btData_Click(object sender, EventArgs e)
        {
            request_data();
        }

        private void btPln_Click(object sender, EventArgs e)
        {
            send_data(cmd_source_pln);
        }

        private void btPltph_Click(object sender, EventArgs e)
        {
            send_data(cmd_source_pltph);
        }

        private void btRefreshChart_Click(object sender, EventArgs e)
        {
            // https://api.thingspeak.com/channels/1476079/charts/1?yaxismin=-2&yaxismax=2&days=1&height=0&width=0
        }

        private void label32_Click(object sender, EventArgs e)
        {
            if(lbProtection.Text == "OFF")
            {
                send_data(cmd_protec_on);
            }
            else
            {
                send_data(cmd_protec_off);
            }
        }

        private void fmMain_Shown(object sender, EventArgs e)
        {
            cbSerial_refresh();
        }

        private void cbSerial_DropDown(object sender, EventArgs e)
        {
            cbSerial_refresh();
        }
    }
}
