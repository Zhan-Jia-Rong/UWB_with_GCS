using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Threading;
using System.Text.RegularExpressions;
using System.IO.Ports;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;


using System.Collections;
using System.Net;
using System.Net.Sockets;
using System.Diagnostics;


using MQTTnet;
using MQTTnet.Adapter;
using MQTTnet.Client;
using MQTTnet.Client.Connecting;
using MQTTnet.Client.Disconnecting;
using MQTTnet.Client.Options;
using MQTTnet.Extensions.ManagedClient;
using MQTTnet.Client.Receiving;
using MQTTnet.Diagnostics;
using MQTTnet.Protocol;
using MQTTnet.Server;
using System.Threading.Tasks;

using System.Numerics; //use vector                             => reference->add reference-> System.Numerics
using System.Web.Script.Serialization;  // use Json format      => reference->add reference-> System.Web.Extensions
using Newtonsoft.Json.Linq;  //use JObject Parse                => reference->add reference->json.NET .NET 4.0

namespace hello
{
    public partial class Form1 : Form
    {
        
        System.IO.Ports.SerialPort serialport = new System.IO.Ports.SerialPort();
        System.Windows.Forms.Timer _timer;
        float time, timenew;

        int RX_Counter = 0;
        int Log_Counter = 0;
        bool send_flag = false;
        bool log_flag = false;
        bool log_flag_set = false;
        byte[] Response = new byte[1024];
        byte[] TX_Data = new byte[10];
        double PI = 3.141592654;
        double EARTH_RADIUS = 6378.137;
        int pos_err = 0;

        public const int I_UWB_LPS_TAG_DATAFRAME0_LENGTH = 128;

        private Boolean receiving;
        private Thread t;
        delegate void Display();

        static DateTime gps_epoch = new DateTime(1980, 1, 6, 0, 0, 0, DateTimeKind.Utc);
        static DateTime unix_epoch = new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc);
        static DateTime my_start = DateTime.UtcNow;
        //static uint GPS_LEAPSECONDS_MILLIS = 18000;

        string pathFile = @"C:\UWBtest\UWBtest2.xlsx";
        Excel.Application excelApp;
        Excel._Workbook wBook;
        Excel._Worksheet wSheet;

        //graphing map
        int m_Width_add = 0;
        const int m_Scale_diameter = 15;
        Bitmap m_image;

        //MQTT tag message
        public string payload;
        public string payloadText;
        string pos_tag;
        byte[] tag_byte = new byte[1024];
        double[] tag0_pos = new double[3];
        double[] tag1_pos = new double[3];

        //MQTT anchor message
        string pos_anchor;
        byte[] anchor_byte = new byte[1024];
        double[] anchor0_pos = new double[2];
        double[] anchor1_pos = new double[2];
        double[] anchor2_pos = new double[2];
        double[] anchor3_pos = new double[2];

        //MQTT only can deliver one messages once,this parameter is used to identify what is the messages type (0=tag,1=anchor,2=disconnect).
        int MQTT_messages_type = 0;



        // Cache font instead of recreating font objects each time we paint.
        private Font fnt = new Font("Arial", 10);

        static class Constants
        {
            public const int MEMBER = 3;            //Flying roboter number
            public const int Tolerant_xy = 10;
            public const int Tolerant_alt = 15;
            public const int Circle_number = 24;
            public const double Radian = 6.28318530717958;
            public const double X_positive_limit = 2;
            public const double Y_positive_limit = 2;
            public const double Z_positive_limit = 2.5;
            public const double X_negative_limit = -2;
            public const double Y_negative_limit = -2;
            public const double Z_negative_limit = 0;
        }

        public class Global
        {
            public static int Counter = 0;

            public static int Navi_counter = 0;
            public static int Number_flag = 0;
            public static int Armed_flag = 0;
            public static int Takeoff_flag = 0;
            public static int Navi_flag = 0;
            public static int Arrive_counter_flag = 0;
            public static int Arrive_counter = 0;
            public static int Reapeat_flag = 0;

            public static int[] buffer = new int[25];
            // 0 |      1     |     2     |       3     |    4     |  5   |  6 |    7   |  8   |  9   | 10   | 11   | 12   | 13   | 14   | 15   | 16   |  17  |  18  |      19      |      20      |     21      |     22      |23|24 |                 
            //|# |Number_flag |Armed_flag |Takeoff_flag |Navi_flag |x_t_h |x_t_l |y_t_h |y_t_l |z_t_h |z_t_l |x_c_h |x_c_l |y_c_h |y_c_l |z_c_h |z_c_l |yaw_h |yaw_l |pitch_speed_h |pitch_speed_l |roll_speed_h |roll_speed_l |% |\n |

            public static double[] rb_x = new double[Constants.MEMBER];
            public static double[] rb_y = new double[Constants.MEMBER];
            public static double[] rb_z = new double[Constants.MEMBER];
            public static double[] rb_yaw = new double[Constants.MEMBER];
            public static double[] rb_x_last = new double[Constants.MEMBER];
            public static double[] rb_y_last = new double[Constants.MEMBER];
            public static double[] rb_x_speed = new double[Constants.MEMBER];
            public static double[] rb_y_speed = new double[Constants.MEMBER];
        }

        

        public Form1()
        {
            InitializeComponent();
            //MqttServerAsync();      //Set MQTT server and load message.
            MqttClientAsync();  //Connect to NTU server and load message.
            foreach (string com in System.IO.Ports.SerialPort.GetPortNames())
            {
                comboBox1.Items.Add(com);
            }
            // Dock the PictureBox to the form and set its background to white.
            //pictureBox1.Dock = DockStyle.Fill;
            //pictureBox1.BackColor = Color.White;
            // Connect the Paint event of the PictureBox to the event handler method.
            pictureBox1.Paint += new System.Windows.Forms.PaintEventHandler(this.pictureBox1_Paint);

            // Add the PictureBox control to the Form.
            this.Controls.Add(pictureBox1);
        }

        public class DroneData
        {
            public IPEndPoint ep;
            public float lastPosX = 0;
            public float lastPosY = 0;
            public float lastPosZ = 0;
            public DateTime lastTime = DateTime.MaxValue;
            public int lost_count = 0;
            public int gps_skip_count = 0;
            public DroneData(string ip, int port)
            {
                ep = new IPEndPoint(IPAddress.Parse(ip), port);
            }
        }

        private static MAVLink.MavlinkParse mavlinkParse = new MAVLink.MavlinkParse();
        private static Socket mavSock = new Socket(AddressFamily.InterNetwork, SocketType.Dgram, ProtocolType.Udp);
        private static Dictionary<string, DroneData> drones = new Dictionary<string, DroneData>(5);
        private static Stopwatch stopwatch;

        public void Main()
        {
            stopwatch = new Stopwatch();
            stopwatch.Start();

            //drones.Add("bebop2", new DroneData("192.168.42.1", 20000));
            drones.Add("bebop2", new DroneData("127.0.0.1", 20000));

            MAVLink.mavlink_system_time_t cmd = new MAVLink.mavlink_system_time_t();
            cmd.time_boot_ms = 0;
            cmd.time_unix_usec = (ulong)((DateTime.UtcNow - unix_epoch).TotalMilliseconds * 1000);
            byte[] pkt = mavlinkParse.GenerateMAVLinkPacket20(MAVLink.MAVLINK_MSG_ID.SYSTEM_TIME, cmd);
            foreach (KeyValuePair<string, DroneData> drone in drones)
            {
                mavSock.SendTo(pkt, drone.Value.ep);
            }
            send_flag = true;

        }
        /*
        public void processFrameData()
        {
            MAVLink.mavlink_att_pos_mocap_t att_pos = new MAVLink.mavlink_att_pos_mocap_t();
            att_pos.time_usec = (ulong)(((DateTime.UtcNow - unix_epoch).TotalMilliseconds - 10) * 1000);
            att_pos.x = Anchor_y; //north Anchor_y
            att_pos.y = Anchor_x; //east Anchor_x
            att_pos.z = Anchor_z; //down
            //att_pos.q = new float[4] { rbData.qw, rbData.qx, rbData.qz, -rbData.qy };

            DroneData drone = drones;
            drone.lost_count = 0;
            byte[] pkt;
            pkt = mavlinkParse.GenerateMAVLinkPacket20(MAVLink.MAVLINK_MSG_ID.ATT_POS_MOCAP, att_pos);
            mavSock.SendTo(pkt, drones.ep);
        }*/
        

        private void btnDate_Click(object sender, EventArgs e)
        {
            send_flag = false;
        }

        double get_distance(double lat1, double lng1, double lat2, double lng2)
        {
            double radLat1 = lat1 * PI / 180.0;   //角度1˚ = π / 180
            double radLat2 = lat2 * PI / 180.0;   //角度1˚ = π / 180
            double a = radLat1 - radLat2;//緯度之差
            double b = lng1 * PI / 180.0 - lng2 * PI / 180.0;  //經度之差
            double dst = 2 * Math.Asin((Math.Sqrt(Math.Pow(Math.Sin(a / 2), 2) + Math.Cos(radLat1) * Math.Cos(radLat2) * Math.Pow(Math.Sin(b / 2), 2))));
            dst = dst * EARTH_RADIUS;
            dst = Math.Round(dst * 10000) / 10000;
            return dst;
        }

        int get_angle(double lat1, double lng1, double lat2, double lng2)
        {
            double x = lat1 - lat2;//t d
            double y = lng1 - lng2;//z y
            int angle = -1;
            if (y == 0 && x > 0) angle = 0;
            if (y == 0 && x < 0) angle = 180;
            if (x == 0 && y > 0) angle = 90;
            if (x == 0 && y < 0) angle = 270;
            if (angle == -1)
            {
                double dislat = get_distance(lat1, lng2, lat2, lng2);
                double dislng = get_distance(lat2, lng1, lat2, lng2);
                if (x > 0 && y > 0) angle = (int) (Math.Atan2(dislng, dislat) / PI * 180);
                if (x < 0 && y > 0) angle = (int) (Math.Atan2(dislat, dislng) / PI * 180 + 90);
                if (x < 0 && y < 0) angle = (int) (Math.Atan2(dislng, dislat) / PI * 180 + 180);
                if (x > 0 && y < 0) angle = (int) (Math.Atan2(dislat, dislng) / PI * 180 + 270);
            }
            label32.Text = Convert.ToString(angle);
            return angle;
        }


        private void Double2Pixel(double dx, double dy, out int px, out int py)
        {
            double picBoxWidth_mid = pictureBox1.Size.Width / 2;
            double picBoxHeight_mid = pictureBox1.Size.Height / 2;
            double picBoxHeight = pictureBox1.Size.Height;
            double scale = 0;
            dx = dx * 100;
            dy = dy * 100;
            scale = (float)m_Width_add / 100;

            //px = (int)(picBoxWidth_mid + dx * scale - m_Scale_diameter / 2);
            //py = (int)(picBoxHeight_mid - dy * scale - m_Scale_diameter / 2);
            px = (int)(0 + dx * scale - m_Scale_diameter / 2);
            py = (int)(picBoxHeight - dy * scale - m_Scale_diameter / 2);
        }


        //*****************************************Graphics***************************************//
        private void pictureBox1_Paint (object sender, PaintEventArgs e)
        {
            int picBoxWidth = pictureBox1.Size.Width;
            int picBoxHeight = pictureBox1.Size.Height;
            //Console.WriteLine($"{picBoxHeight}__{picBoxWidth}");
            //int picBoxWidth_mid = pictureBox1.Size.Width/2;
            //int picBoxHeight_mid = pictureBox1.Size.Height/2;
            int picBoxWidth_mid = 0;
            int count = 0;
            int widthMax = 19; //8
            int hieighMax = 19; //5
            int width = 0;
            //int px = 0;
            //int py = 0;

            Graphics objGraphic = e.Graphics; //**請注意這一行** 
            Font drawFont = new Font("Arial", 8);
            SolidBrush darwBrush = new SolidBrush(Color.Black);
            Pen pen = new Pen(Color.Black, 10);
            Pen pen_line = new Pen(Color.Black);
            SolidBrush darwBrush_circle = new SolidBrush(Color.Red);

            if (hieighMax >= widthMax)
            {
                m_Width_add = picBoxHeight / hieighMax;
            }
            else
            {
                m_Width_add = picBoxWidth / widthMax;
            }

            for (width = 0; width < picBoxWidth;) //畫格線 起點左下(0.0)
            {
                objGraphic.DrawLine(pen_line, picBoxWidth_mid + width, 0, picBoxWidth_mid + width, picBoxHeight);
                objGraphic.DrawString(count.ToString("F00"), drawFont, darwBrush, picBoxWidth_mid + width, picBoxHeight - 20);    // draw NUMBER toString Parameter F
                objGraphic.DrawLine(pen_line, 0, picBoxHeight - width, picBoxWidth, picBoxHeight - width); // Draw text(y > 0)
                objGraphic.DrawString((count * (-1)).ToString("F00"), drawFont, darwBrush, picBoxWidth_mid, picBoxHeight - width);
                count = count + 1;
                width = width + m_Width_add;
            }

            //Draw Frame
            objGraphic.DrawLine(pen, 0, 0, picBoxWidth, 0);
            objGraphic.DrawLine(pen, 0, 0, 0, picBoxHeight);
            objGraphic.DrawLine(pen, picBoxWidth, 0, picBoxWidth, picBoxHeight);
            objGraphic.DrawLine(pen, 0, picBoxHeight, picBoxWidth, picBoxHeight);
        }

        private void draw_point()
        {
            int i, j;

            double[] x = new double[Constants.MEMBER];
            double[] y = new double[Constants.MEMBER];
            double[] z = new double[Constants.MEMBER];
            string[] str_x = new string[Constants.MEMBER];
            string[] str_y = new string[Constants.MEMBER];

            SolidBrush darwBrushInfo = new SolidBrush(Color.Black);
            SolidBrush darwBrushPos1 = new SolidBrush(Color.Red);
            SolidBrush darwBrushPos2 = new SolidBrush(Color.Green);
            SolidBrush darwBrushPos3 = new SolidBrush(Color.Blue);

            SolidBrush darwBrushWaypoint1 = new SolidBrush(Color.Pink);
            SolidBrush darwBrushWaypoint2 = new SolidBrush(Color.GreenYellow);
            SolidBrush darwBrushWaypoint3 = new SolidBrush(Color.RoyalBlue);
            Font drawFont = new Font("Arial", 8);
            m_image = new Bitmap(pictureBox1.Width, pictureBox1.Height);
            for (i = 0; i < Global.Counter; i++)
            {
                //str_x[0] = WaypointView.Items[i].SubItems[1].Text;
                //str_y[0] = WaypointView.Items[i].SubItems[2].Text;
                //str_x[1] = WaypointView.Items[i].SubItems[4].Text;
                //str_y[1] = WaypointView.Items[i].SubItems[5].Text;
                //str_x[2] = WaypointView.Items[i].SubItems[7].Text;
                //str_y[2] = WaypointView.Items[i].SubItems[8].Text;
                for (j = 0; j < Constants.MEMBER; j++)
                {
                    x[j] = Convert.ToDouble(str_x[j]);
                    y[j] = Convert.ToDouble(str_y[j]);
                    drawString(Graphics.FromImage(m_image), drawFont, darwBrushInfo, x[j], y[j], i);        // draw coordinate information
                    if (j == 0)
                    {
                        drawPoint(Graphics.FromImage(m_image), darwBrushWaypoint1, 5, 5);                 // draw waypoint
                    }
                    else if (j == 1)
                    {
                        drawPoint(Graphics.FromImage(m_image), darwBrushWaypoint2, x[j], y[j]);                 // draw waypoint
                    }
                    else
                    {
                        drawPoint(Graphics.FromImage(m_image), darwBrushWaypoint3, x[j], y[j]);                 // draw waypoint
                    }
                }
            }
            for (i = 0; i < Constants.MEMBER; i++)
            {
                x[i] = Convert.ToDouble(Global.rb_x[i]);
                y[i] = Convert.ToDouble(Global.rb_y[i]);
                z[i] = Global.rb_z[i] * 100 - 1;
                drawString(Graphics.FromImage(m_image), drawFont, darwBrushInfo, x[i], y[i], z[i]);         // draw flying roboter coordinate information
                if (i == 0)
                {
                    drawPoint(Graphics.FromImage(m_image), darwBrushPos1, Global.rb_x[i], Global.rb_y[i]);      // draw flying roboter position
                }
                else if (i == 1)
                {
                    drawPoint(Graphics.FromImage(m_image), darwBrushPos2, Global.rb_x[i], Global.rb_y[i]);
                }
                else
                {
                    drawPoint(Graphics.FromImage(m_image), darwBrushPos3, Global.rb_x[i], Global.rb_y[i]);
                }
            }
            pictureBox1.Image = m_image;
        }

        private void drawPoint(Graphics g, SolidBrush darwBrush, double x, double y)
        {
            int px, py;
            Double2Pixel(x, y, out px, out py);                                         // Converter double to pixel
            g.FillEllipse(darwBrush, px, py, m_Scale_diameter, m_Scale_diameter);       // Draw point  
        }
        private void drawString(Graphics g, Font drawFont, SolidBrush darwBrush, double x, double y, double count)
        {
            int px, py;
            count = count + 1;
            Double2Pixel(x, y, out px, out py);                                         // Converter double to pixel
            g.FillEllipse(darwBrush, px, py, m_Scale_diameter, m_Scale_diameter);       // Draw point  
            g.DrawString(count.ToString("F01"), drawFont, darwBrush, px + 20, py - 10);                  // Draw text(y < 0)
        }

        private void btnQuit_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
        private void use_str_Spilt()
        {
            string[] words;
            char[] delimiterChars = { ',', ':', '"', '[', ']' };
            int UWB_tag_id;
            int UWB_tag_grounp;
            int UWB_anchor_grounp;

            if (pos_tag != null && MQTT_messages_type == 0)
            {
                words = new[] { pos_tag };   // string to string[]
                words = pos_tag.Split(delimiterChars);   // String split according to "delimiterChars"
                UWB_tag_id = int.Parse(words[9]);
                UWB_tag_grounp = int.Parse(words[13]);
                if (UWB_tag_id == 4)
                {
                    tag0_pos[0] = double.Parse(words[18]);   // string to double
                    tag0_pos[1] = double.Parse(words[19]);
                    tag0_pos[2] = double.Parse(words[20]);

                }
                if (UWB_tag_id == 5)
                {
                    tag1_pos[0] = double.Parse(words[18]);   // string to double
                    tag1_pos[1] = double.Parse(words[19]);
                    tag1_pos[2] = double.Parse(words[20]);

                }
                if (pos_anchor != null && MQTT_messages_type == 1)
                {
                    //Console.WriteLine(pos_anchor);
                    words = new[] { pos_anchor };   // string to string[]
                                                    //Console.WriteLine("spilt:");
                    words = pos_anchor.Split(delimiterChars);   // String split according to "delimiterChars"

                    UWB_anchor_grounp = int.Parse(words[9]);

                    //foreach(var word in words){Console.WriteLine($"< {word} >");}

                    anchor0_pos[0] = double.Parse(words[15]);   // string to double
                    anchor0_pos[1] = double.Parse(words[16]);
                    anchor1_pos[0] = double.Parse(words[19]);
                    anchor1_pos[1] = double.Parse(words[20]);
                    anchor2_pos[0] = double.Parse(words[23]);
                    anchor2_pos[1] = double.Parse(words[24]);
                    anchor3_pos[0] = double.Parse(words[27]);
                    anchor3_pos[1] = double.Parse(words[28]);
                }
                if (MQTT_messages_type != 2)
                {
                    Display d = new Display(DisplayText);
                    this.Invoke(d);
                }
                else
                {

                }
            }
        }
        private void use_json_format_spilt()
        {
            int UWB_tag_id;
            int UWB_tag_grounp;
            int UWB_anchor_grounp;
            if (payloadText.Contains("pan_id"))  //Make sure the "JObject.Parse()" has not error.
            {
                var getResult = JObject.Parse(payloadText); //Json format spilt.

                //Console.WriteLine(payload);
                UWB_tag_id = (int)getResult["euid"];
                UWB_tag_grounp = (int)getResult["pan_id"];

                if (UWB_tag_id == 01)
                {
                    tag0_pos[0] = (double)getResult["global_pos"][0];
                    tag0_pos[1] = (double)getResult["global_pos"][1];
                    tag0_pos[2] = (double)getResult["global_pos"][2];
                    Console.WriteLine((string)getResult["global_pos"][0]);
                }
                if (UWB_tag_id == 02)
                {
                    tag1_pos[0] = (double)getResult["global_pos"][0];
                    tag1_pos[1] = (double)getResult["global_pos"][1];
                    tag1_pos[2] = (double)getResult["global_pos"][2];
                    Console.WriteLine((string)getResult["global_pos"][0]);
                }
                /*
                if ((string)getResult["command"] == "anchor")
                {
                    UWB_anchor_grounp = (int)getResult["group"];

                    anchor0_pos[0] = (double)getResult["anchor_pos"][0][0];   // string to double
                    anchor0_pos[1] = (double)getResult["anchor_pos"][0][1];

                    anchor1_pos[0] = (double)getResult["anchor_pos"][1][0];
                    anchor1_pos[1] = (double)getResult["anchor_pos"][1][1];

                    anchor2_pos[0] = (double)getResult["anchor_pos"][2][0];
                    anchor2_pos[1] = (double)getResult["anchor_pos"][2][1];

                    anchor3_pos[0] = (double)getResult["anchor_pos"][3][0];
                    anchor3_pos[1] = (double)getResult["anchor_pos"][3][1];
                }
                */
                if (payloadText.Contains("reconnected") !=true)
                {
                    Display d = new Display(DisplayText);
                    this.Invoke(d);
                }
                else
                {

                }
            }
        }
        private void DisplayText()
        {
            /*
            label9.Text = buffer[0].ToString("X2");
            label12.Text = Convert.ToString(buffer.Length);
            totalLength = totalLength + buffer.Length;
            label10.Text = totalLength.ToString();
            */
            check_uwb1();
        }
        private void DoReceive()
        {
            while (receiving)
            {
                //use_str_Spilt();
                use_json_format_spilt();
            }
        }

        float theta = 35 + 90;//角度值 4f->64
        double uwb_radians = 0;
        double uwb_angle = 0;
        double uwb_angle_display = 0;
        double uwb_radians_display = 0;
        double radian = 0;
        private System.Windows.Forms.Timer timer1;
        private void check_uwb1()
        {
            label9.Text = "UWB OK";
            label26.Text = Convert.ToString(tag0_pos[0]);
            label27.Text = Convert.ToString(tag0_pos[1]);
            label33.Text = Convert.ToString(tag1_pos[0]);
            label34.Text = Convert.ToString(tag1_pos[1]);
            uwb_radians = Math.Atan2(tag1_pos[1] - tag0_pos[1], tag1_pos[0] - tag0_pos[0]); //set tag0_pos close (0.0)
            uwb_radians = uwb_radians - radian;   
            uwb_angle = -uwb_radians * (180 / Math.PI);   
            if (uwb_angle < 0) uwb_angle = uwb_angle + 360;  // (-uwb_radians) mean ( 0~180 and 0~-180 )to 0~360

            label32.Text = Convert.ToString((int)uwb_angle);
            label28.Text = Convert.ToString((float)uwb_radians);

            Font drawFont = new Font("Arial", 8);
            m_image = new Bitmap(pictureBox1.Width, pictureBox1.Height);

            SolidBrush darwBrushPos2 = new SolidBrush(Color.Blue);
            SolidBrush darwBrushPos3 = new SolidBrush(Color.Red);
            drawPoint(Graphics.FromImage(m_image), darwBrushPos3, tag0_pos[0], tag0_pos[1]);
            drawPoint(Graphics.FromImage(m_image), darwBrushPos2, tag1_pos[0], tag1_pos[1]);
            pictureBox1.Image = m_image;

            if (send_flag)
            {
                long cur_ms = stopwatch.ElapsedMilliseconds;
                MAVLink.mavlink_vision_position_estimate_t vision_position = new MAVLink.mavlink_vision_position_estimate_t();

                vision_position.usec = (ulong)(cur_ms * 1000);

                vision_position.x = (float)(tag1_pos[0] * Math.Cos(radian) + tag1_pos[1] * Math.Sin(radian)); //north for use compass
                vision_position.y = (float)-((-tag1_pos[0] * Math.Sin(radian)) + tag1_pos[1] * Math.Cos(radian)); //east  for use compass

                vision_position.yaw = (float)-uwb_radians;  //In NED frame the Z-axis points down to represent positive. so we need add negative.

                label29.Text = Convert.ToString(vision_position.x);
                label30.Text = Convert.ToString(vision_position.y);

                DroneData drone = drones["bebop2"];
                drone.lost_count = 0;

                byte[] pkt;
                if (tag1_pos[0] > -10 && tag1_pos[0] < 40 && tag1_pos[1] > -10 && tag1_pos[1] < 40) // limit
                {
                    pkt = mavlinkParse.GenerateMAVLinkPacket20(MAVLink.MAVLINK_MSG_ID.VISION_POSITION_ESTIMATE, vision_position);
                    mavSock.SendTo(pkt, drone.ep);
                }
                else
                {
                    pos_err = pos_err + 1;
                    label36.Text = Convert.ToString(pos_err);
                }
            }

        }


        private void button2_Click(object sender, EventArgs e)
        {
            receiving = true;
            t = new Thread(DoReceive);
            t.IsBackground = true;
            t.Start();
        }


        private void Form1_Load(object sender, EventArgs e)
        {
        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {
        }

        private void button4_Click(object sender, EventArgs e)
        {
        }

        private void button5_Click(object sender, EventArgs e)
        {
            log_flag = true;

            excelApp = new Excel.Application();    // 開啟一個新的應用程式
            //excelApp.Visible = true;             // 讓Excel文件可見
            excelApp.DisplayAlerts = false;        // 停用警告訊息
            excelApp.Workbooks.Add(Type.Missing);  // 加入新的活頁簿
            wBook = excelApp.Workbooks[1];         // 引用第一個活頁簿
            wBook.Activate();                      // 設定活頁簿焦點

            timer1 = new System.Windows.Forms.Timer();
            timer1.Interval = 500; // 設定計時器的速度
            timer1.Tick += new EventHandler(timer1_Tick);
            timer1.Start();
            time = 0;
            timenew = 0;



        }

        private void timer1_Tick(object sender, EventArgs e)
        {

            label37.Text = (int)timenew / 10 + " seconds";

            if (log_flag)
            {
                Log_Counter = Log_Counter + 1;
                if (tag1_pos[1] > 0)
                {
                    try
                    {
                        wSheet = (Excel._Worksheet)wBook.Worksheets[1];   // 引用第一個工作表
                        wSheet.Name = "UWB Sensor Value Log";   // 命名工作表的名稱
                        wSheet.Activate();  // 設定工作表焦點  

                        excelApp.Cells[Log_Counter, 1] = timenew / 10; //att_pos.x
                        excelApp.Cells[Log_Counter, 2] = tag1_pos[0]; //att_pos.y
                        excelApp.Cells[Log_Counter, 3] = tag1_pos[1]; //att_pos.y
                                                                   //excelApp.Cells[Log_Counter, 3] = Anchor_z; //att_pos.z
                                                                   //excelApp.Cells[Log_Counter, 3] = Anchor_x1; //att_pos.x
                                                                   //excelApp.Cells[Log_Counter, 4] = Anchor_y1; //att_pos.y
                                                                   //excelApp.Cells[Log_Counter, 5] = uwb_angle;
                                                                   //excelApp.Cells[Log_Counter, 6] = (Anchor_x + Anchor_x1)/2;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("產生報表時出錯！" + Environment.NewLine + ex.Message);
                    }
                }
                if (log_flag_set)
                {
                    try
                    {
                        //另存活頁簿
                        wBook.SaveAs(pathFile, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                        Console.WriteLine("儲存文件於 " + Environment.NewLine + pathFile);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("儲存檔案出錯，檔案可能正在使用" + Environment.NewLine + ex.Message);
                    }

                    wBook.Close(false, Type.Missing, Type.Missing);   //關閉活頁簿
                    excelApp.Quit();  //關閉Excel
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);  //釋放Excel資源
                    wBook = null;
                    wSheet = null;
                    excelApp = null;
                    GC.Collect();
                    Console.Read();
                    log_flag = false;
                    log_flag_set = false;
                    Log_Counter = 0;
                }
            }
            timenew = timenew + 5;
        }



        private void button1_Click(object sender, EventArgs e)
        {
            log_flag_set = true;
        }


        private void button3_Click_1(object sender, EventArgs e)
        {
            textBox1.Clear();
            draw_point();
            //pictureBox1.Refresh();
        }

        //DroneData drone = drones["bebop2"];

        private void button4_Click_1(object sender, EventArgs e)
        {
            MAVLink.mavlink_set_mode_t cmd = new MAVLink.mavlink_set_mode_t();
            cmd.base_mode = (byte)MAVLink.MAV_MODE_FLAG.CUSTOM_MODE_ENABLED;
            cmd.custom_mode = 9;
            cmd.target_system = 1;
            DroneData drone = drones["bebop2"];
            byte[] pkt = mavlinkParse.GenerateMAVLinkPacket20(MAVLink.MAVLINK_MSG_ID.SET_MODE, cmd);
            mavSock.SendTo(pkt, drone.ep);
        }

        private void button7_Click(object sender, EventArgs e)
        {
            MAVLink.mavlink_set_mode_t cmd = new MAVLink.mavlink_set_mode_t();
            cmd.base_mode = (byte)MAVLink.MAV_MODE_FLAG.CUSTOM_MODE_ENABLED;
            cmd.custom_mode = 3;
            cmd.target_system = 1;
            DroneData drone = drones["bebop2"];
            byte[] pkt = mavlinkParse.GenerateMAVLinkPacket20(MAVLink.MAVLINK_MSG_ID.SET_MODE, cmd);
            mavSock.SendTo(pkt, drone.ep);
        }

        private void button8_Click(object sender, EventArgs e)
        {
            MAVLink.mavlink_command_long_t cmd = new MAVLink.mavlink_command_long_t();
            cmd.command = (ushort)MAVLink.MAV_CMD.COMPONENT_ARM_DISARM;
            cmd.target_system = 1;
            cmd.param1 = 0;
            cmd.param2 = 21196;
            byte[] pkt = mavlinkParse.GenerateMAVLinkPacket20(MAVLink.MAVLINK_MSG_ID.COMMAND_LONG, cmd);
            DroneData drone = drones["bebop2"];
            mavSock.SendTo(pkt, drone.ep);
        }

        private void button9_Click(object sender, EventArgs e)
        {
            MAVLink.mavlink_command_long_t cmd = new MAVLink.mavlink_command_long_t();
            cmd.command = (ushort)MAVLink.MAV_CMD.TAKEOFF;
            cmd.target_system = 1;
            //cmd.target_component = 250;
            cmd.param7 = 1.0f;
            DroneData drone = drones["bebop2"];
            byte[] pkt = mavlinkParse.GenerateMAVLinkPacket20(MAVLink.MAVLINK_MSG_ID.COMMAND_LONG, cmd);
            mavSock.SendTo(pkt, drone.ep);
        }

        private void button10_Click(object sender, EventArgs e)
        {
            MAVLink.mavlink_set_position_target_local_ned_t cmd = new MAVLink.mavlink_set_position_target_local_ned_t();
            cmd.target_system = 1;
            cmd.coordinate_frame = (byte)MAVLink.MAV_FRAME.LOCAL_NED;
            cmd.type_mask = 0xff8;

            theta = Convert.ToInt32(textBox2.Text) + 90;
            //label32.Text = Convert.ToString(theta);
            radian = theta * Math.PI / 180;//轉換弧度值

            float waypoint1_x = 0.0f + float.Parse(textBox3.Text);
            float waypoint1_y = 0.0f + float.Parse(textBox4.Text);
            cmd.x = (float)(waypoint1_x * Math.Cos(radian) + waypoint1_y * Math.Sin(radian)); //north
            cmd.y = (float)-((-waypoint1_x * Math.Sin(radian)) + waypoint1_y * Math.Cos(radian)); //east
            //cmd.x = waypoint1_x;
            //cmd.y = -waypoint1_y;
            //label33.Text = Convert.ToString(cmd.x);
            //label34.Text = Convert.ToString(cmd.y);

            cmd.z = -1.5f;
            DroneData drone = drones["bebop2"];
            byte[] pkt = mavlinkParse.GenerateMAVLinkPacket20(MAVLink.MAVLINK_MSG_ID.SET_POSITION_TARGET_LOCAL_NED, cmd);
            mavSock.SendTo(pkt, drone.ep);
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {

        }

        private void button14_Click(object sender, EventArgs e)
        {
            MAVLink.mavlink_command_long_t cmd = new MAVLink.mavlink_command_long_t();
            cmd.command = (ushort)MAVLink.MAV_CMD.COMPONENT_ARM_DISARM;
            cmd.target_system = 1;
            cmd.param1 = 1;
            DroneData drone = drones["bebop2"];
            byte[] pkt = mavlinkParse.GenerateMAVLinkPacket20(MAVLink.MAVLINK_MSG_ID.COMMAND_LONG, cmd);
            mavSock.SendTo(pkt, drone.ep);
        }

        private void button15_Click(object sender, EventArgs e)
        {
            MAVLink.mavlink_set_position_target_local_ned_t cmd = new MAVLink.mavlink_set_position_target_local_ned_t();
            cmd.target_system = 1;
            cmd.coordinate_frame = (byte)MAVLink.MAV_FRAME.LOCAL_NED;
            cmd.type_mask = 0xff8;

            //theta = Convert.ToInt32(textBox2.Text) + 90;
            //label32.Text = Convert.ToString(theta);
            //radian = theta * Math.PI / 180;//轉換弧度值

            float waypoint1_x = 0.0f + float.Parse(textBox7.Text);
            float waypoint1_y = 0.0f + float.Parse(textBox8.Text);
            //cmd.x = (float)(waypoint1_x * Math.Cos(radian) + waypoint1_y * Math.Sin(radian)); //north
            //cmd.y = (float)-((-waypoint1_x * Math.Sin(radian)) + waypoint1_y * Math.Cos(radian)); //east
            cmd.x = waypoint1_x;
            cmd.y = -waypoint1_y;
            //label33.Text = Convert.ToString(cmd.x);
            //label34.Text = Convert.ToString(cmd.y);

            cmd.z = -1.6f;
            DroneData drone = drones["bebop2"];
            byte[] pkt = mavlinkParse.GenerateMAVLinkPacket20(MAVLink.MAVLINK_MSG_ID.SET_POSITION_TARGET_LOCAL_NED, cmd);
            mavSock.SendTo(pkt, drone.ep);
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }

        private void button16_Click(object sender, EventArgs e)
        {
            //int anglee7 = get_angle(24.7732410, 121.0452770, 39.945908, 116.906084);
            //int anglee7 = get_angle(24.7732240, 121.0452490, 24.7736180, 121.045638);
            //int anglee7 = get_angle(24.7735950, 121.0458900, 24.7734779, 21.0457900); //(A0,A3) 
            int anglee7 = get_angle(24.7734668, 121.0457839, 24.7735950, 121.0458900); //(A3,A0)
            //int anglee7 = get_angle(24.785207, 121.056423, 24.785240, 121.056395); //中正橋
        }

        private void button11_Click(object sender, EventArgs e)
        {
            MAVLink.mavlink_set_mode_t cmd = new MAVLink.mavlink_set_mode_t();
            cmd.base_mode = (byte)MAVLink.MAV_MODE_FLAG.CUSTOM_MODE_ENABLED;
            cmd.custom_mode = 4;
            cmd.target_system = 1; //depend on drone systemID
            DroneData drone = drones["bebop2"];
            byte[] pkt = mavlinkParse.GenerateMAVLinkPacket20(MAVLink.MAVLINK_MSG_ID.SET_MODE, cmd);
            mavSock.SendTo(pkt, drone.ep);
        }

        private void button12_Click(object sender, EventArgs e)
        {
            MAVLink.mavlink_set_gps_global_origin_t cmd = new MAVLink.mavlink_set_gps_global_origin_t();
            cmd.target_system = 1;
            if (double.Parse(textBox7.Text) > 0)
            {
                cmd.latitude = (int)(double.Parse(textBox7.Text) * 10000000);
                cmd.longitude = (int)(double.Parse(textBox8.Text) * 10000000);
            }
            else
            {
                //cmd.latitude = (int)(24.7734501 * 10000000);
                //cmd.longitude = (int)(121.0459265 * 10000000);

                //cmd.latitude = (int)(24.7978601 * 10000000); //高鐵橋
                //cmd.longitude = (int)(121.0375601 * 10000000);

                //cmd.latitude = (int)(23.7780360 * 10000000); //竹山橋
                //cmd.longitude = (int)(120.7097101 * 10000000);

                cmd.latitude = (int)(23.8314592 * 10000000); //自強橋
                cmd.longitude = (int)(120.3983619 * 10000000);

            }


            cmd.altitude = 100;
            DroneData drone = drones["bebop2"];
            byte[] pkt = mavlinkParse.GenerateMAVLinkPacket20(MAVLink.MAVLINK_MSG_ID.SET_GPS_GLOBAL_ORIGIN, cmd);
            mavSock.SendTo(pkt, drone.ep);
        }

        private void button13_Click(object sender, EventArgs e)
        {
            MAVLink.mavlink_set_position_target_local_ned_t cmd = new MAVLink.mavlink_set_position_target_local_ned_t();
            cmd.target_system = 1;
            cmd.coordinate_frame = (byte)MAVLink.MAV_FRAME.LOCAL_NED;
            cmd.type_mask = 0xff8;

            theta = Convert.ToInt32(textBox2.Text) + 90;
            //label32.Text = Convert.ToString(theta);
            radian = theta * Math.PI / 180;//轉換弧度值

            float waypoint2_x = 0.0f + float.Parse(textBox5.Text);
            float waypoint2_y = 0.0f + float.Parse(textBox6.Text);
            cmd.x = (float)(waypoint2_x * Math.Cos(radian) + waypoint2_y * Math.Sin(radian)); //north
            cmd.y = (float)-((-waypoint2_x * Math.Sin(radian)) + waypoint2_y * Math.Cos(radian)); //east
            //cmd.x = waypoint2_x;
            //cmd.y = -waypoint2_y;
            //label33.Text = Convert.ToString(cmd.x);
            //label34.Text = Convert.ToString(cmd.y);

            cmd.z = -1.5f;
            DroneData drone = drones["bebop2"];
            byte[] pkt = mavlinkParse.GenerateMAVLinkPacket20(MAVLink.MAVLINK_MSG_ID.SET_POSITION_TARGET_LOCAL_NED, cmd);
            mavSock.SendTo(pkt, drone.ep);
        }

        private void button6_Click(object sender, EventArgs e)
        {
                Main();
                lblShow.Text = "System start";
        }

        private void button3_Click(object sender, EventArgs e)
        {
        }

        //-------------------------------mqtt set-------------------------------------------------
        private async Task MqttServerAsync()
        {
            
            var optionsBuilder = new MqttServerOptionsBuilder().WithConnectionBacklog(100).WithDefaultEndpointPort(1883);
            var mqttServer = new MqttFactory().CreateMqttServer();
            await mqttServer.StartAsync(optionsBuilder.Build());
            mqttServer.ApplicationMessageReceivedHandler = new MqttApplicationMessageReceivedHandlerDelegate(e =>
            {
                Console.WriteLine($"Client:{e.ClientId} Topic:{e.ApplicationMessage.Topic} Message:{Encoding.UTF8.GetString(e.ApplicationMessage.Payload ?? new byte[0])}");
                payload = Encoding.UTF8.GetString(e.ApplicationMessage.Payload ?? new byte[0]);
                if (payload.Contains("disconnect"))
                {
                    //Console.WriteLine(Encoding.UTF8.GetString(e.ApplicationMessage.Payload ?? new byte[0]));
                    MQTT_messages_type = 2;
                }
                else
                {
                    if (payload.Contains("tag") && payload.Length > 50)
                    {
                        //use string spilt
                        tag_byte = e.ApplicationMessage.Payload ?? new byte[0];
                        pos_tag = Encoding.UTF8.GetString(tag_byte);
                        MQTT_messages_type = 0;
                    }
                    else if (payload.Contains("anchor") && payload.Length > 50)
                    {
                        /*  
                        //use json format spilt
                        var getResult = JObject.Parse(payload);
                        Console.WriteLine(getResult["anchor_pos"][0][0]);
                        Console.WriteLine("---------");
                        double[] anchor_pos_test = new double[2];
                        anchor_pos_test[0] = (double)getResult["anchor_pos"][0][0];
                        */

                        //use string spilt
                        anchor_byte = e.ApplicationMessage.Payload ?? new byte[0];
                        pos_anchor = Encoding.UTF8.GetString(anchor_byte);
                        MQTT_messages_type = 1;
                    }
                }

            });
            mqttServer.ClientConnectedHandler = new MqttServerClientConnectedHandlerDelegate(e =>
            {
                //Console.WriteLine($"Client:{e.ClientId} 已連接！");
            });
            mqttServer.ClientDisconnectedHandler = new MqttServerClientDisconnectedHandlerDelegate(e =>
            {
                //Console.WriteLine($"Client:{e.ClientId}已離線！");
            });
        }

        private async Task MqttClientAsync()
        {
            var factory = new MqttFactory();

            var clientOptions = new MqttClientOptionsBuilder()
                .WithClientId("TestClient")
                .WithTcpServer("140.112.45.232", 1883)
                .Build();

            var client = factory.CreateMqttClient();
            client.ApplicationMessageReceivedHandler = new MqttApplicationMessageReceivedHandlerDelegate(msg =>
            {
                payloadText = Encoding.UTF8.GetString(msg?.ApplicationMessage?.Payload ?? Array.Empty<byte>());

                //Console.WriteLine($"Received msg: {payloadText}");
            });

            /*
            client.UseApplicationMessageReceivedHandler(msg =>
            {
                var payloadText = Encoding.UTF8.GetString(
                    msg?.ApplicationMessage?.Payload ?? Array.Empty<byte>());

                Console.WriteLine($"Received msg: {payloadText}");
            });
            */
            await client.ConnectAsync(clientOptions, CancellationToken.None);

            await client.SubscribeAsync(
                new MqttTopicFilter
                {
                    Topic = "ITRI1_Result"
                    //Topic = "ITRI1_MonitorReport"
                    //Topic = "ITRI1_SensorMessage"
                }
            );
        }
        //-------------------------------mqtt set-------------------------------------------------

    }



}
