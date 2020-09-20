using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO.Ports;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;

namespace RFID
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            System.Windows.Forms.Control.CheckForIllegalCrossThreadCalls = false;//设置该属性 为false
        }


        private void Form1_Load(object sender, EventArgs e)
        {
            RegistryKey keyCom = Registry.LocalMachine.OpenSubKey("Hardware\\DeviceMap\\SerialComm");
            if (keyCom != null)
            {
                string[] sSubKeys = keyCom.GetValueNames();
                comboBox1.Items.Clear();
                foreach (string sName in sSubKeys)
                {
                    string sValue = (string)keyCom.GetValue(sName);
                    comboBox1.Items.Add(sValue);
                }
                if (comboBox1.Items.Count > 0)
                {
                    comboBox1.SelectedIndex = 0;
                    comboBox2.SelectedIndex = 0;
                }
            }
        }

        DataSet myDataSet = new DataSet();

        /// <summary>
        /// 打开串口
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            initSerial();
        }

        /// <summary>
        /// 关闭串口
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                serialPort.Close();     //关闭串口
                comboBox1.Enabled = true;//打开使能
                comboBox2.Enabled = true;
                isOpened = false;
            }
            catch
            {
                MessageBox.Show("串口关闭失败！");
            }
        }

        /// <summary>
        /// 串口初始化
        /// </summary>
        bool isOpened = false;
        private void initSerial()
        {
            if (!isOpened)
            {
                serialPort.PortName = comboBox1.Text;
                serialPort.BaudRate = Convert.ToInt32(comboBox2.Text, 10);
                try
                {
                    serialPort.Open();     //打开串口
                    comboBox1.Enabled = false;//关闭使能
                    comboBox2.Enabled = false;
                    isOpened = true;
                    serialPort.DataReceived += new SerialDataReceivedEventHandler(post_DataReceived);//串口接收处理函数
                }
                catch
                {
                    MessageBox.Show("串口打开失败！");
                }
            }
        }

        /// <summary>
        /// 串口读取函数
        /// </summary>
        string str, str_num = "", str_con="";
        private void post_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            str = "";
            str = serialPort.ReadExisting();//字符串方式读
            if (str.Length == 26)
            {
                str_num = str.Substring(0, 8);
                str_con = str.Substring(8, 16);

                try
                {
                    string strCon = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=RFID.accdb";
                    OleDbConnection myconn = new OleDbConnection(strCon);
                    string strcom = null;
                    strcom = "SELECT 卡号,姓名,员工编号 FROM IC WHERE 卡号= '" + str_num + "'";
                    myconn.Open();
                    OleDbDataAdapter mycommand = new OleDbDataAdapter(strcom, myconn);
                    myDataSet.Clear();
                    mycommand.Fill(myDataSet, "select");
                    if (myDataSet.Tables.Count > 0 && myDataSet.Tables["select"].Rows.Count > 0) //判断用户名是否存在
                    {
                        string st1 = myDataSet.Tables["select"].Rows[0].ItemArray[2].ToString();  //查询到得员工编号
                        if (st1.Equals(str_con))  //员工编号是否正确
                        {
                            string destr = "INSERT INTO Record(卡号,姓名,员工编号,打卡时间) VALUES ('" + str_num + "','" + myDataSet.Tables["select"].Rows[0].ItemArray[1].ToString() + "','" + str_con + "','" + DateTime.Now.ToString() + "')";
                            OleDbCommand dest = new OleDbCommand(destr, myconn);
                            dest.ExecuteNonQuery();
                            serialPort.Write("JT\r\n");
                        }
                        else
                        {
                            serialPort.Write("JF\r\n");
                        }
                    }
                    else
                    {
                        serialPort.Write("JF\r\n");
                    }
                    myDataSet.Clear();
                    myconn.Close();
                }
                catch (Exception x)
                {
                    MessageBox.Show("错误!" + x.ToString());
                }

            }
            else if(str.Length == 10)
            {
                str_num = str.Substring(0, 8);
            }
        }


        /// <summary>
        /// 考勤信息
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void 考勤信息ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form2 f2 = new Form2();
            f2.Show();
        }

        /// <summary>
        /// IC卡开户
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// 
        private void button3_Click(object sender, EventArgs e)
        {
            if (textBox1.Text != "" && textBox2.Text != "")
            {
                DialogResult dr = MessageBox.Show("是否确认写入？", "提示", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (dr == DialogResult.Yes)
                {
                    try
                    {
                        string strCon = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=RFID.accdb";
                        OleDbConnection myconn = new OleDbConnection(strCon);
                        string strcom = "SELECT 卡号,姓名,员工编号 FROM IC WHERE 卡号= '" + str_num + "'";
                        myconn.Open();
                        OleDbDataAdapter mycommand = new OleDbDataAdapter(strcom, myconn);
                        myDataSet.Clear();
                        mycommand.Fill(myDataSet, "select");
                        if (myDataSet.Tables["select"].Rows.Count <= 0)
                        {
                            string destr = "INSERT INTO IC(卡号,姓名,员工编号) VALUES ('" + str_num + "','" + textBox2.Text + "','" + textBox1.Text + "')";
                            OleDbCommand dest = new OleDbCommand(destr, myconn);
                            dest.ExecuteNonQuery();
                            MessageBox.Show("成功！", "提示");
                            string s = "W";
                            s += textBox1.Text;
                            s += "\r\n";
                            serialPort.Write(s);
                        }
                        else
                        {
                            MessageBox.Show("卡号已存在！请注销后重试!", "提示");
                        }
                        myDataSet.Clear();
                        myconn.Close();
                    }
                    catch (Exception x)
                    {
                        MessageBox.Show("错误！" + x.ToString());
                    }
                }
            }
            else
            {
                MessageBox.Show("输入不能为空", "提示");
            }
        }

        /// <summary>
        /// 注销员工IC卡
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button4_Click(object sender, EventArgs e)
        {
            if (textBox1.Text != "")
            {
                DialogResult dr = MessageBox.Show("是否确认注销？", "提示", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (dr == DialogResult.Yes)
                {
                    try
                    {
                        string strCon = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=RFID.accdb";
                        OleDbConnection myconn = new OleDbConnection(strCon);
                        string strcom = "SELECT 卡号,姓名,员工编号 FROM IC WHERE 员工编号= '" + textBox1.Text + "'";
                        myconn.Open();
                        OleDbDataAdapter mycommand = new OleDbDataAdapter(strcom, myconn);
                        myDataSet.Clear();
                        mycommand.Fill(myDataSet, "select");
                        if (myDataSet.Tables["select"].Rows.Count > 0)
                        {
                            string destr = "DELETE FROM IC WHERE 员工编号 = '" + textBox1.Text + "'";
                            OleDbCommand dest = new OleDbCommand(destr, myconn);
                            dest.ExecuteNonQuery();
                            MessageBox.Show("注销成功！", "提示");
                        }
                        else
                        {
                            MessageBox.Show("卡号不存在！", "提示");
                        }
                        myconn.Close();
                    }
                    catch (Exception x)
                    {
                        MessageBox.Show("错误！" + x.ToString());
                    }
                }
            }
            else
            {
                MessageBox.Show("请输入员工编号", "提示");
            }
        }

    }
}
