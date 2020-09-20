using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;

namespace RFID
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }

        DataSet myDataSet = new DataSet();

        private void Form2_Load(object sender, EventArgs e)
        {
            string strCon = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=RFID.accdb";
            OleDbConnection myconn = new OleDbConnection(strCon);
            string strcom = "SELECT * FROM Record";
            myconn.Open();
            OleDbDataAdapter mycommand = new OleDbDataAdapter(strcom, myconn);
            myDataSet.Clear();
            mycommand.Fill(myDataSet, "Record");
            myconn.Close();
            dataGridView1.Columns[2].DefaultCellStyle.Format = "yyyy/MM/dd HH:mm:ss";
            this.dataGridView1.DataSource = myDataSet.Tables["Record"];
        }

        /// <summary>
        /// 查询指定员工编号的打卡记录
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text != "" && textBox1.Text.Length == 16)
            {
                string strCon = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=RFID.accdb";
                OleDbConnection myconn = new OleDbConnection(strCon);
                string strcom = "SELECT * FROM Record WHERE 员工编号 = '" + textBox1.Text + "'";
                myconn.Open();
                OleDbDataAdapter mycommand = new OleDbDataAdapter(strcom, myconn);
                myDataSet.Clear();
                mycommand.Fill(myDataSet, "Record");
                myconn.Close();
                if (myDataSet.Tables.Count > 0 && myDataSet.Tables["Record"].Rows.Count > 0) //判断用户名是否存在
                {
                    this.dataGridView1.DataSource = myDataSet.Tables["Record"];
                }
                else
                {
                    MessageBox.Show("查询不存在！", "提示");
                }
            }
            else
            {
                MessageBox.Show("请输入正确的员工编号！", "提示");
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            int l = this.dataGridView1.CurrentRow.Index;

            DialogResult dr = MessageBox.Show("确定删除该条记录?", "提示", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (dr == DialogResult.Yes)
            {
                try
                {
                    string st = myDataSet.Tables["Record"].Rows[l].ItemArray[2].ToString();
                    string strCon = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=RFID.accdb";
                    OleDbConnection myconn = new OleDbConnection(strCon);
                    myconn.Open();
                    string strcom = "SELECT * FROM Record";
                    string destr = "DELETE FROM Record WHERE 打卡时间 = #" + st + "#";
                    OleDbCommand dest = new OleDbCommand(destr, myconn);
                    dest.ExecuteNonQuery();
                    MessageBox.Show("删除成功！", "提示");

                    OleDbDataAdapter mycommand = new OleDbDataAdapter(strcom, myconn);
                    myDataSet.Clear();
                    mycommand.Fill(myDataSet, "Record");
                    myconn.Close();
                }
                catch (Exception x)
                {
                    MessageBox.Show("错误！" + x.ToString());
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            DialogResult dr = MessageBox.Show("是否清空所有考勤记录?", "谨慎操作！", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (dr == DialogResult.Yes)
            {
                try
                {
                    string strCon = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=RFID.accdb";
                    OleDbConnection myconn = new OleDbConnection(strCon);
                    myconn.Open();
                    string strcom = "SELECT * FROM Record";
                    string destr = "DELETE FROM Record";
                    OleDbCommand dest = new OleDbCommand(destr, myconn);
                    dest.ExecuteNonQuery();
                    MessageBox.Show("删除成功！", "提示");

                    OleDbDataAdapter mycommand = new OleDbDataAdapter(strcom, myconn);
                    myDataSet.Clear();
                    mycommand.Fill(myDataSet, "Record");
                    myconn.Close();
                }
                catch (Exception x)
                {
                    MessageBox.Show("错误！" + x.ToString());
                }
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            string strCon = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=RFID.accdb";
            OleDbConnection myconn = new OleDbConnection(strCon);
            string strcom = "SELECT * FROM Record";
            myconn.Open();
            OleDbDataAdapter mycommand = new OleDbDataAdapter(strcom, myconn);
            myDataSet.Clear();
            mycommand.Fill(myDataSet, "Record");
            myconn.Close();
            this.dataGridView1.DataSource = myDataSet.Tables["Record"];
        }
    }
}
