using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.NetworkInformation;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;

namespace rdpRemote
{

    public partial class Form1 : Form
    {
        public class ConnectionInfo
        {
            public string 地址 { get; set; }
            public string 用户名 { get; set; }
            public string 密码 { get; set; }
            public string 文件名 { get; set; }
            public string 备注 { get; set; }
        }
        public Form1()
        {
            InitializeComponent();
        }
        string configPath = System.Windows.Forms.Application.StartupPath + @"\Resources\ServersConfig.xml";
        private void Form1_Load(object sender, EventArgs e)
        {
            InitDataGridView();
           

        }

        private void InitDataGridView()
        {
            List<ConnectionInfo> result = new List<ConnectionInfo>();

            XmlOp xml = new XmlOp(configPath);
            var connectInfo = new ConnectionInfo();
            var root = xml.GetXmlRoot();
            var nodeList = root.GetEnumerator();
            while (nodeList.MoveNext())
            {
                XmlAttributeCollection temp = ((XmlNode)nodeList.Current).Attributes;
                var subNode = temp.GetEnumerator();
                var conNode = new ConnectionInfo();
                while (subNode.MoveNext())
                {
                    XmlAttribute attribute = (XmlAttribute)subNode.Current;
                    if (attribute.Name == "ip")
                    {
                        conNode.地址 = attribute.Value;
                    }
                }
                conNode.用户名 = xml.GetXmlNodeValue(@"/servers/server[@ip='" + conNode.地址 + "']/UserName");
                conNode.文件名 = xml.GetXmlNodeValue(@"/servers/server[@ip='" + conNode.地址 + "']/FileName");
                conNode.密码 = xml.GetXmlNodeValue(@"/servers/server[@ip='" + conNode.地址 + "']/PassWord");
                conNode.备注 = xml.GetXmlNodeValue(@"/servers/server[@ip='" + conNode.地址 + "']/Remark");
                result.Add(conNode);
            }
            dataGridView1.DataSource = result;
            dataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            DataGridViewButtonColumn connectBtn = new DataGridViewButtonColumn();
            connectBtn.HeaderText = "远程登录";
            connectBtn.Text = "远程登录";
            connectBtn.UseColumnTextForButtonValue = true;
            connectBtn.AutoSizeMode =
                DataGridViewAutoSizeColumnMode.AllCells;
            connectBtn.FlatStyle = FlatStyle.Standard;
            connectBtn.CellTemplate.Style.BackColor = Color.Honeydew;
            dataGridView1.Columns.Add(connectBtn);

            DataGridViewButtonColumn closeBtn = new DataGridViewButtonColumn();
            closeBtn.HeaderText = "远程关机";
            closeBtn.Text = "远程关机";
            closeBtn.UseColumnTextForButtonValue = true;
            closeBtn.AutoSizeMode =
                DataGridViewAutoSizeColumnMode.AllCells;
            closeBtn.FlatStyle = FlatStyle.Standard;
            closeBtn.CellTemplate.Style.BackColor = Color.Honeydew;
            dataGridView1.Columns.Add(closeBtn);


            DataGridViewButtonColumn pingBtn = new DataGridViewButtonColumn();
            pingBtn.HeaderText = "Ping";
            pingBtn.Text = "Ping";
            pingBtn.UseColumnTextForButtonValue = true;
            pingBtn.AutoSizeMode =
                DataGridViewAutoSizeColumnMode.AllCells;
            pingBtn.FlatStyle = FlatStyle.Standard;
            pingBtn.CellTemplate.Style.BackColor = Color.Honeydew;
            dataGridView1.Columns.Add(pingBtn);

            DataGridViewImageColumn imageCell = new DataGridViewImageColumn();
            imageCell.HeaderText = "状态";
            imageCell.Image = Image.FromFile(Application.StartupPath + @"\Resources\red.jpg");
            imageCell.AutoSizeMode =
                DataGridViewAutoSizeColumnMode.AllCells;
            dataGridView1.Columns.Add(imageCell);
        }
        private void DoPing(int rowIndex)
        {
            //tim.Enabled = false;
            Application.DoEvents();
            var con = new ConnectionInfo();
            con.地址 = this.dataGridView1.Rows[rowIndex].Cells[0].Value.ToString();
            con.用户名 = this.dataGridView1.Rows[rowIndex].Cells[1].Value.ToString();
            con.密码 = this.dataGridView1.Rows[rowIndex].Cells[2].Value.ToString();
            con.文件名 = this.dataGridView1.Rows[rowIndex].Cells[3].Value.ToString();
            con.备注 = this.dataGridView1.Rows[rowIndex].Cells[4].Value.ToString();
            if (Ping(con.地址))
            {
                dataGridView1.Rows[rowIndex].Cells[8].Value = Image.FromFile(Application.StartupPath + @"\Resources\green.jpg");
            }
            else
            {
                dataGridView1.Rows[rowIndex].Cells[8].Value = Image.FromFile(Application.StartupPath + @"\Resources\red.jpg");
            }
            // tim.Enabled = true;
        }
        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 5) //远程登录
            {
                var con = new ConnectionInfo();
                con.地址 = this.dataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString();
                con.用户名 = this.dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString();
                con.密码 = this.dataGridView1.Rows[e.RowIndex].Cells[2].Value.ToString();
                con.文件名 = this.dataGridView1.Rows[e.RowIndex].Cells[3].Value.ToString();
                con.备注 = this.dataGridView1.Rows[e.RowIndex].Cells[4].Value.ToString();
                Connection(con);
            }
            if (e.ColumnIndex == 6) //关机
            {
                var con = new ConnectionInfo();
                con.地址 = this.dataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString();
                con.用户名 = this.dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString();
                con.密码 = this.dataGridView1.Rows[e.RowIndex].Cells[2].Value.ToString();
                con.文件名 = this.dataGridView1.Rows[e.RowIndex].Cells[3].Value.ToString();
                con.备注 = this.dataGridView1.Rows[e.RowIndex].Cells[4].Value.ToString();
                DialogResult dr = MessageBox.Show("确定要关闭吗?", "关闭系统", MessageBoxButtons.YesNo);
                if (dr == DialogResult.Yes)
                {
                    RemoteClose(con);
                }
            }
            if (e.ColumnIndex == 7) //Ping
            {
                var con = new ConnectionInfo();
                con.地址 = this.dataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString();
                con.用户名 = this.dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString();
                con.密码 = this.dataGridView1.Rows[e.RowIndex].Cells[2].Value.ToString();
                con.文件名 = this.dataGridView1.Rows[e.RowIndex].Cells[3].Value.ToString();
                con.备注 = this.dataGridView1.Rows[e.RowIndex].Cells[4].Value.ToString();
                if (Ping(con.地址))
                {
                    dataGridView1.Rows[e.RowIndex].Cells[8].Value = Image.FromFile(Application.StartupPath + @"\Resources\green.jpg");
                }
                else
                {
                    dataGridView1.Rows[e.RowIndex].Cells[8].Value = Image.FromFile(Application.StartupPath + @"\Resources\red.jpg");
                }
            }
        }

        #region FUNC
        private bool Ping(string address)
        {
            Ping p1 = new Ping();

            PingReply reply = p1.Send(address);

            StringBuilder sbuilder;
            if (reply.Status == IPStatus.Success)
            {
                return true;

            }
            else
            {
                return false;
            }
        }
        private void RemoteClose(ConnectionInfo con)
        {
            List<string> cmdLst = new List<string>();
            var cmd0 = "net use \\\\" + con.地址 + "\\ipc$ \"" + con.密码 + "\" /user:\"" + con.用户名 + "\"";
            var cmd1 = "shutdown -s -m \\\\" + con.地址 + " -t 10";
            cmdLst.Add(cmd0);
            cmdLst.Add(cmd1);
            ExeCommand(cmdLst.ToArray());
        }
        private void RemoteLogin(String cmd)
        {
            Process p = new Process();
            p.StartInfo.FileName = "cmd.exe";
            p.StartInfo.UseShellExecute = false;
            p.StartInfo.RedirectStandardInput = true;
            p.StartInfo.RedirectStandardOutput = true;
            p.StartInfo.RedirectStandardError = true;
            p.StartInfo.CreateNoWindow = true;
            p.Start();
            p.StandardInput.WriteLine(cmd);
            p.StandardInput.WriteLine("exit");
        }
        private void Connection(ConnectionInfo con)
        {
            string address = con.地址;
            string username = con.用户名;
            string password = con.密码;
            string filename = con.文件名;
            var TemplateStr = rdpRemote.Properties.Resources.TemplateRDP;//获取RDP模板字符串
            //用DataProtection加密密码,并转化成二进制字符串
            var pwstr = BitConverter.ToString(DataProtection.ProtectData(Encoding.Unicode.GetBytes(password), ""));
            pwstr = pwstr.Replace("-", "");
            //替换模板里面的关键字符串,生成当前的drp字符串
            var NewStr = TemplateStr.Replace("{#address}", address).Replace("{#username}", username).Replace("{#password}", pwstr);
            //将drp保存到文件，并放在程序目录下，等待使用
            StreamWriter sw = new StreamWriter(filename);
            sw.Write(NewStr);
            sw.Close();
            //利用CMD命令调用MSTSC
            RemoteLogin("mstsc " + filename);
        }
        public static string ExeCommand(string[] commandTexts)
        {
            Process p = new Process();
            p.StartInfo.FileName = "cmd.exe";
            p.StartInfo.UseShellExecute = false;
            p.StartInfo.RedirectStandardInput = true;
            p.StartInfo.RedirectStandardOutput = true;
            p.StartInfo.RedirectStandardError = true;
            p.StartInfo.CreateNoWindow = true;
            string strOutput = null;
            try
            {
                p.Start();
                foreach (string item in commandTexts)
                {
                    p.StandardInput.WriteLine(item);
                }
                p.StandardInput.WriteLine("exit");
                strOutput = p.StandardOutput.ReadToEnd();
                p.WaitForExit();
                p.Close();
            }
            catch (Exception e)
            {
                strOutput = e.Message;
            }
            return strOutput;
        }
        #endregion

        private void toolStripButton1_Click_1(object sender, EventArgs e)
        {
            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                DoPing(i);

            }
        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("notepad.exe", configPath);
        }

        private void toolStripButton3_Click(object sender, EventArgs e)
        {
            dataGridView1.Columns.Clear();
            InitDataGridView();
        }
    }
}
