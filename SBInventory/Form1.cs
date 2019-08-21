using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Data.Odbc;
using System.Net;
using System.Net.Mail;
using System.Web;
using System.Threading;
using System.Data.SqlClient;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.IO;
using System.Runtime.InteropServices;
namespace SBInventory
{
    
    public partial class Form1 : Form
    {
        public const int WM_NCLBUTTONDOWN = 0xA1;
        public const int HT_CAPTION = 0x2;
        [DllImportAttribute("user32.dll")]
        public static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);
        [DllImportAttribute("user32.dll")]
        public static extern bool ReleaseCapture();
        string dataGridSelectedLootNumber = "";
        string dataGridSelectedDID = "";
        string dataGridSelectedDSB = "";
        string dataGridSelectedEID = "";
        string dataGridSelectedIsByHand = "";
        DataTable dtDSB;
        string update_form = "";
        public static string sb_name_static = "";
        bool QueryFlag = false;
        public Form1()
        {
            this.WindowState = FormWindowState.Maximized;
            InitializeComponent();
        }
        OleDbConnection conn;
        private void exit_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void entry_Click(object sender, EventArgs e)
        {
            //tableLayoutPanel4.Visible = true;
        }

        private void tableLayoutPanel4_Paint(object sender, PaintEventArgs e)
        {

        }

        private void label6_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
           
        }
        
        private void Form1_Load(object sender, EventArgs e)
        {
            label31.Text = Form1.sb_name_static;
            comboBox8.SelectedIndex = 2;
            comboBox9.SelectedIndex = 2;

            var DBPath = Application.StartupPath + "\\db.mdb";
            conn = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=db.mdb;Jet OLEDB:Database Password=qlty;");
            conn.Open();
            using (DataTable dt = new DataTable())
            {              
                String sql = "SELECT SBDSB FROM tblSBDSB WHERE active_stat=1";
                using (OleDbDataAdapter adapter = new OleDbDataAdapter(sql, conn))
                {
                    adapter.Fill(dt);
                    dtDSB = dt;
                }
                
               
                foreach (DataRow row in dt.Rows)
                {
                    ComboboxItem item = new ComboboxItem();
                    item.Text = row[0].ToString();
                    item.Value = row[0].ToString();
                    comboBox3.Items.Add(item);
                    //comboBox19.Items.Add(item);
                    comboBox1.Items.Add(item);
                    comboBox11.Items.Add(item);
                }
                              
               // comboBox19.SelectedIndex = 0;
                //comboBox3.SelectedIndex = 0;
                //comboBox2.SelectedIndex = 0;
               // comboBox4.SelectedIndex = 0;
            }
            using (DataTable dt = new DataTable())
            {

                String sql = "SELECT DISTINCT LotNo FROM tblSBDispatch";
                using (OleDbDataAdapter adapter = new OleDbDataAdapter(sql, conn))
                {
                    adapter.Fill(dt);
                }

                foreach (DataRow row in dt.Rows)
                {
                    ComboboxItem item = new ComboboxItem();
                    item.Text = row[0].ToString();
                    item.Value = row[0].ToString();
                    comboBox7.Items.Add(item);
                }

                ComboboxItem item1 = new ComboboxItem();
                item1.Text = "No Loot/Reference Number";
                item1.Value = "No Loot/Reference Number";
                comboBox7.Items.Add(item1);
            }
            conn.Close();
            comboBox3.SelectedIndex = 0;
            comboBox3.SelectedIndex = 0;
            comboBox1.SelectedIndex = 0;
            comboBox11.SelectedIndex = 0;
            textBox1.Focus();

            string startDate = dateTimePicker5.Value.ToString("d");
            string endDate = dateTimePicker6.Value.ToString("d");

            using (DataTable dt = new DataTable())
            {

                String sql = "SELECT DISTINCT LotNo FROM tblSBDispatch WHERE DRDate BETWEEN #" + startDate + "# AND #" + endDate + "#";
                using (OleDbDataAdapter adapter = new OleDbDataAdapter(sql, conn))
                {
                    adapter.Fill(dt);
                }
                comboBox7.Items.Clear();
                foreach (DataRow row in dt.Rows)
                {
                    ComboboxItem item = new ComboboxItem();
                    item.Text = row[0].ToString();
                    item.Value = row[0].ToString();
                    comboBox7.Items.Add(item);
                }
                ComboboxItem item1 = new ComboboxItem();
                item1.Text = "No Loot/Reference Number";
                item1.Value = "No Loot/Reference Number";
                comboBox7.Items.Add(item1);
            }
            if (QueryFlag)
            {
                tabControl1.TabPages.RemoveByKey("tabPage2");
                tabControl1.TabPages.RemoveByKey("tabPage3");
                tabControl1.TabPages.RemoveByKey("tabPage6");
                tabControl1.TabPages.RemoveByKey("tabPage7");
                tabControl1.TabPages.RemoveByKey("tabPage8");
            }           
            
        }

        private void button3_Click(object sender, EventArgs e)
        {
            showQueryResult();
        }
        public void showQueryResult()
        {
            string eid = textBox4.Text;
            if (eid.Length == 15)
            {
                using (DataTable dt = new DataTable())
                {

                    string sql_search = @"SELECT 
                                            DRDate AS `DATE`,
                                            FORMAT(DRTime, 'Long Time') AS `TIME`,
                                            SBDSB,
                                            IIF(STS='RD','Re-Dispatch', 'New Dispatch') AS STATUS,
                                            LotNo AS `LOT No`,
                                            IIF(isByHand=1,'By Hand', 'By Post') AS Dispatch,
                                            DID,EID 
                                          FROM 
                                            tblSBDispatch 
                                          WHERE EID='" + eid + "'";
                    using (OleDbDataAdapter adapter = new OleDbDataAdapter(sql_search, conn))
                    {
                        adapter.Fill(dt);
                    }

                    dataGridView1.DataSource = dt;
                    dataGridView1.Columns["DID"].Visible = false;
                    dataGridView1.Columns["EID"].Visible = false;
                    if (dataGridView1.ColumnCount<=8)
                    {
                        DataGridViewImageColumn img = new DataGridViewImageColumn();
                        System.Drawing.Image image = SBInventory.Properties.Resources.edit;
                        img.Image = image;
                        img.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                        dataGridView1.Columns.Add(img);
                        img.HeaderText = "Action";
                        img.Name = "img";
                    }
                    
                   
 
                }

                using (DataTable dt = new DataTable())
                {
                    string sql_search = @"SELECT 
                                            DRDate AS `DATE`,
                                            FORMAT(DRTime, 'Long Time') AS `TIME`,
                                            SBDSB,
                                            switch(
                                                STS ='P', 'POSITIVE',
                                                STS ='N', 'NEGATIVE'
                                              ) AS RESULT,
                                            RID 
                                        FROM 
                                            tblSBReceive 
                                        WHERE 
                                            EID='" + eid + "'";
                    using (OleDbDataAdapter adapter = new OleDbDataAdapter(sql_search, conn))
                    {
                        adapter.Fill(dt);
                    }
                    
                    dataGridView2.DataSource = dt;
                    dataGridView2.Columns["RID"].Visible = false;

                    if (dataGridView2.ColumnCount <= 5)
                    {
                        DataGridViewImageColumn img = new DataGridViewImageColumn();
                        System.Drawing.Image image = SBInventory.Properties.Resources.edit;
                        img.Image = image;
                        dataGridView2.Columns.Add(img);
                        img.HeaderText = "Action";
                        img.Name = "img";
                    }
                    for (int i = 0; i < dataGridView2.Rows.Count; i++)
                    {
                        if (dataGridView2.Rows[i].Cells["RESULT"].FormattedValue.ToString().Trim() == "NEGATIVE")
                        {
                            dataGridView2.Rows[i].Cells["RESULT"].Style = new DataGridViewCellStyle { ForeColor = Color.Red };
                        }
                    }
                }
            }
            else { MessageBox.Show("ENROLLMENT ID NOT GIVEN PROPERLY"); }
        }
        public static string exportSQL = "";
        private void button4_Click(object sender, EventArgs e)
        {
            label10.Text = "Database is Uploading Please wait";
            button4.Enabled = false;
            Thread thread1 = new Thread(new ThreadStart(sendMail));
            thread1.Start();
            
        }

        public string exportData(string tableName)
        {
            
            return null;
        }
        public void sendMail()
        {
            try
            {
                // /send email
                string Subject = "DATABASE BACKUP OF SB INVENTORY SYSTEM (" + sb_name_static + ")";
                string ToEmail = "munkadirbd@gmail.com";
                string SMTPUser = "munkadirbd@gmail.com", SMTPPassword = "halarpo123";
                List<string> sql = new List<string>();
                using (DataTable dt = new DataTable())
                {

                    string sql_search = "SELECT * FROM tblSBDispatch";
                    using (OleDbDataAdapter adapter = new OleDbDataAdapter(sql_search, conn))
                    {
                        adapter.Fill(dt);
                    }

                    foreach (DataRow row in dt.Rows)
                    {
                       string[] drdateArry=  row[1].ToString().Split(' ');
                       string[] drtimeArry = row[2].ToString().Split(' ');
                       sql.Add(@"INSERT INTO tblSBDispatch (DID, DRDate, DRTime, EID, Remarks, SBDSB,STS,LotNo,isByHand)
                                    VALUES (" + "'" + row[0].ToString() + "','" + drdateArry[0] + "','" + drtimeArry[1] + " " + drtimeArry[2] + "','" + row[3].ToString() + "','" + row[4].ToString() + "','" + row[5].ToString() + "','" + row[6].ToString() + "','" + row[7].ToString() + "'" + "," + row[8].ToString() + ")");
                       // exportSQL = exportSQL + row[0].ToString() + " " + row[1].ToString() + " " + row[2].ToString() + " " + row[3].ToString() + " " + row[4].ToString();
                    }

                    using (DataTable dtReceive = new DataTable())
                    {

                        string sqlReceive = "SELECT * FROM tblSBReceive";
                        using (OleDbDataAdapter adapterReceive = new OleDbDataAdapter(sqlReceive, conn))
                        {
                            adapterReceive.Fill(dtReceive);
                        }

                        foreach (DataRow rowReceive in dtReceive.Rows)
                        {

                            string[] drdateArryReceive = rowReceive[1].ToString().Split(' ');
                            string[] drtimeArryReceive = rowReceive[2].ToString().Split(' ');
                            sql.Add(@"INSERT INTO tblSBReceive (RID, DRDate, DRTime, EID, Remarks, SBDSB,STS)
                                    VALUES (" + "'" + rowReceive[0].ToString() + "','" + drdateArryReceive[0] + "','" + drtimeArryReceive[1] + " " + drtimeArryReceive[2] + "','" + rowReceive[3].ToString() + "','" + rowReceive[4].ToString() + "','" + rowReceive[5].ToString() + "','" + rowReceive[6].ToString() + "',')");
                            // exportSQL = exportSQL + row[0].ToString() + " " + row[1].ToString() + " " + row[2].ToString() + " " + row[3].ToString() + " " + row[4].ToString();
                        }
                        SmtpClient smtp = new SmtpClient();
                        smtp.UseDefaultCredentials = true;
                        smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
                        smtp.EnableSsl = true;
                        MailMessage mail = new MailMessage();
                        mail.To.Add(ToEmail);
                        mail.From = new MailAddress(SMTPUser);
                        mail.Subject = Subject;
                        mail.Body = string.Join(";", sql.ToArray());
                        mail.IsBodyHtml = true;
                        //if you are using your smtp server, then change your host like "smtp.yourdomain.com"
                        smtp.Host = "smtp.gmail.com";
                        //change your port for your host
                        smtp.Port = 25; //or you can also use port# 587
                        smtp.Credentials = new System.Net.NetworkCredential(SMTPUser, SMTPPassword);
                        //smtp.Host = "smtp.gmail.com";               

                        smtp.Send(mail);
                        
                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("Upload FAILED.\nPlease Check Your Internet Connection");
            }
            finally { label10.Text = "Database Sucessfully Uploaded";button4.Enabled = true; }
        }

        private void button5_Click(object sender, EventArgs e)
        {
        }

        private void button6_Click(object sender, EventArgs e)
        {
            
        }

        private void button7_Click(object sender, EventArgs e)
        {
        }

        private void tableLayoutPanel10_Paint(object sender, PaintEventArgs e)
        {

        }

        private void tableLayoutPanel5_Paint(object sender, PaintEventArgs e)
        {

        }
     
        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
           
        }

        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            
        }

        private void button8_Click(object sender, EventArgs e)
        {            
            showQueryResult();   
        }

        private void Form1_MouseDown(object sender, MouseEventArgs e)
        {
           /* if (e.Button == MouseButtons.Left)
            {
                ReleaseCapture();
                SendMessage(Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0);
            }*/
        }

        private void tableLayoutPanel1_MouseDown(object sender, MouseEventArgs e)
        {
           /* if (e.Button == MouseButtons.Left)
            {
                ReleaseCapture();
                SendMessage(Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0);
            }*/
        }

        private void button10_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        private void tableLayoutPanel13_MouseDown(object sender, MouseEventArgs e)
        {
            /*if (e.Button == MouseButtons.Left)
            {
                ReleaseCapture();
                SendMessage(Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0);
            }*/
        }

        private void button9_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void button11_Click(object sender, EventArgs e)
        {
            if (this.WindowState == FormWindowState.Maximized)
            {
                this.WindowState = FormWindowState.Normal;
            }
            else 
            { 
                this.WindowState = FormWindowState.Maximized;
            }
        }

        private void tableLayoutPanel2_Paint(object sender, PaintEventArgs e)
        {

        }

        private void button10_Click_1(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            
            string date = DateTime.Today.ToString("MM-dd-yyyy");
            string time = string.Format("{0:hh:mm:ss tt}", DateTime.Now);
            string loot = textBox2.Text.Trim();
            string eid = textBox1.Text.Trim();
            string SBDSB = comboBox3.Text.Trim();
            bool dispatchFlag = false;
            int totalDispatch = 0;
            int totalReceive = 0;
            if (eid.Length == 15)
            {
                using (DataTable dt = new DataTable())
                {
                    string sql_search = "SELECT count(*) FROM tblSBDispatch WHERE EID='" + eid + "'";
                    using (OleDbDataAdapter adapter = new OleDbDataAdapter(sql_search, conn))
                    {
                        adapter.Fill(dt);
                    }


                    foreach (DataRow row in dt.Rows)
                    {
                        if (row[0].ToString() == "0")
                        {
                            totalDispatch = 0;
                            string isByHand = "0";
                            if (checkBox1.Checked)
                            {
                                isByHand = "1";
                            }
                            string sql_insrt = @"INSERT INTO tblSBDispatch (DRDate, DRTime, EID, Remarks, SBDSB, STS,LotNo,isByHand)
                            VALUES ('" + date + "','" + time + "','" + eid + "','Dispatch','" + SBDSB + "','ND','" + loot + "'," + isByHand + ")";
                            conn.Open();
                            using (OleDbCommand cmd = new OleDbCommand(sql_insrt, conn))
                            {
                                cmd.ExecuteNonQuery();
                            }
                            textBox1.Text = "";
                            //textBox2.Text = "";
                            label27.ForeColor = Color.DarkGreen;
                            if (comboBox7.FindString(loot) == -1)
                            {
                                ComboboxItem iteam = new ComboboxItem();
                                iteam.Text = loot;
                                iteam.Value = loot;
                                comboBox7.Items.Add(iteam);
                            }
                            label27.Text = "Dispatch Sucessful\nEnrollment ID: " + eid + "\nLoot: " + loot + "\nDate: " + date + "\nTime: " + time + "\n" + SBDSB;
                            checkBox1.Checked = false;
                            conn.Close();
                        }
                        else
                        {
                            //MessageBox.Show("not new");
                            totalDispatch = Convert.ToInt32(row[0].ToString());
                            using (DataTable dtstat = new DataTable())
                            {
                                string sql_search_stat = "SELECT STS FROM tblSBReceive WHERE EID='" + eid + "'";
                                using (OleDbDataAdapter adapter = new OleDbDataAdapter(sql_search_stat, conn))
                                {
                                    adapter.Fill(dtstat);
                                }
                                totalReceive = dtstat.Rows.Count;
                                foreach (DataRow rowStat in dtstat.Rows)
                                {
                                    if (rowStat[0].ToString() == "N")
                                    {
                                        dispatchFlag = true;
                                    }
                                    else { dispatchFlag = false; }
                                }

                                if (dispatchFlag)
                                {
                                    string isByHand = "0";
                                    if (checkBox1.Checked)
                                    {
                                        isByHand = "1";
                                    }
                                    string sql_insrt = @"INSERT INTO tblSBDispatch (DRDate, DRTime, EID, Remarks, SBDSB, STS,LotNo,isByHand)
                                                     VALUES ('" + date + "','" + time + "','" + eid + "','Dispatch','" + SBDSB + "','RD','" + loot + "'," + isByHand + ")";
                                    conn.Open();
                                    using (OleDbCommand cmd = new OleDbCommand(sql_insrt, conn))
                                    {
                                        cmd.ExecuteNonQuery();
                                    }
                                    textBox1.Text = "";
                                    //textBox2.Text = "";
                                    label27.ForeColor = Color.BlueViolet;
                                    checkBox1.Checked = false;
                                    label27.Text = "Redispatch Sucessful\nEnrollment ID: " + eid + "\nLoot: " + loot + "\nDate: " + date + "\nTime: " + time + "\n" + SBDSB;
                                    if (comboBox7.FindString(loot) == -1)
                                    {
                                        ComboboxItem iteam = new ComboboxItem();
                                        iteam.Text = loot;
                                        iteam.Value = loot;
                                        comboBox7.Items.Add(iteam);
                                    }
                                    conn.Close();
                                }
                                else
                                {
                                    label27.ForeColor = Color.DarkRed;
                                    label27.Text = "Already Ditchpatched";
                                }
                            }
                        }

                    }
                }
                textBox1.Text = "";
            }
            else { MessageBox.Show("ENROLLMENT ID NOT GIVEN PROPERLY"); }
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            string date = DateTime.Today.ToString("MM-dd-yyyy");
            string time = string.Format("{0:hh:mm:ss tt}", DateTime.Now);
            string eid = textBox3.Text;
            string sts = "P";
            bool dispatchFlag = true;
            string SBDSB = "";
            int totalDispatch = 0;
            int totalReceive = 0;
            if (eid.Length == 15)
            {
                using (DataTable dt = new DataTable())
                {
                    string sql_search = "SELECT count(*) FROM tblSBDispatch WHERE EID='" + eid + "'";
                    using (OleDbDataAdapter adapter = new OleDbDataAdapter(sql_search, conn))
                    {
                        adapter.Fill(dt);
                    }


                    foreach (DataRow row in dt.Rows)
                    {
                        if (row[0].ToString() == "0")
                        {
                            totalDispatch = 0;
                            label28.ForeColor = Color.DarkRed;
                            label28.Text = "Not Dispatched Yet";
                        }
                        else
                        {
                            using (DataTable dtstatSBDSB = new DataTable())
                            {
                                string sql_search_stat_SBDBS = "SELECT SBDSB FROM tblSBDispatch WHERE EID='" + eid + "'";
                                using (OleDbDataAdapter adapter = new OleDbDataAdapter(sql_search_stat_SBDBS, conn))
                                {
                                    adapter.Fill(dtstatSBDSB);
                                }
                                totalDispatch = Convert.ToInt32(row[0].ToString());
                                foreach (DataRow rowStatSBDSB in dtstatSBDSB.Rows)
                                {
                                    SBDSB = rowStatSBDSB[0].ToString();
                                }
                            }

                            using (DataTable dtstat = new DataTable())
                            {
                                string sql_search_stat = "SELECT STS FROM tblSBReceive WHERE EID='" + eid + "'";
                                using (OleDbDataAdapter adapter = new OleDbDataAdapter(sql_search_stat, conn))
                                {
                                    adapter.Fill(dtstat);
                                }
                                totalReceive = dtstat.Rows.Count;
                                foreach (DataRow rowStat in dtstat.Rows)
                                {
                                    if (rowStat[0].ToString() == "N")
                                    {
                                        dispatchFlag = true;
                                    }
                                    else { dispatchFlag = false; }
                                }

                                if (dispatchFlag && (totalDispatch > totalReceive))
                                {
                                    string sql_insrt = @"INSERT INTO tblSBReceive (DRDate, DRTime, EID, Remarks, SBDSB, STS)
                                                     VALUES ('" + date + "','" + time + "','" + eid + "','Receive','" + SBDSB + "','" + sts + "')";
                                    conn.Open();
                                    //MessageBox.Show(sql_insrt);
                                    using (OleDbCommand cmd = new OleDbCommand(sql_insrt, conn))
                                    {
                                        cmd.ExecuteNonQuery();
                                    }
                                    conn.Close();
                                    label28.ForeColor = Color.DarkGreen;
                                    label28.Text = "Receive Successful (POSITIVE)\nEnrollment ID: " + eid + "\nDate: " + date + "\nTime: " + time + "\n" + SBDSB;
                                    conn.Close();
                                }
                                else
                                {
                                    label28.ForeColor = Color.DarkRed;
                                    label28.Text = "Already Received";
                                }
                            }
                        }

                    }

                }
                textBox3.Text = "";
            }
            else { MessageBox.Show("ENROLLMENT ID NOT GIVEN PROPERLY"); }
        }

        private void button7_Click_1(object sender, EventArgs e)
        {

            string date = DateTime.Today.ToString("MM-dd-yyyy");
            string time = string.Format("{0:hh:mm:ss tt}", DateTime.Now);
            string eid = textBox5.Text;
            string sts = "N";
            bool dispatchFlag = true;
            string SBDSB = "";
            int totalDispatch = 0;
            int totalReceive = 0;
            if (eid.Length == 15)
            {
                using (DataTable dt = new DataTable())
                {
                    string sql_search = "SELECT count(*) FROM tblSBDispatch WHERE EID='" + eid + "'";
                    using (OleDbDataAdapter adapter = new OleDbDataAdapter(sql_search, conn))
                    {
                        adapter.Fill(dt);
                    }


                    foreach (DataRow row in dt.Rows)
                    {
                        if (row[0].ToString() == "0")
                        {
                            totalDispatch = 0;
                            label29.ForeColor = Color.DarkRed;
                            label29.Text ="Not Dispatched Yet";
                        }
                        else
                        {
                            using (DataTable dtstatSBDSB = new DataTable())
                            {
                                totalDispatch = Convert.ToInt32(row[0].ToString());
                                string sql_search_stat_SBDBS = "SELECT SBDSB FROM tblSBDispatch WHERE EID='" + eid + "'";
                                using (OleDbDataAdapter adapter = new OleDbDataAdapter(sql_search_stat_SBDBS, conn))
                                {
                                    adapter.Fill(dtstatSBDSB);
                                }
                                foreach (DataRow rowStatSBDSB in dtstatSBDSB.Rows)
                                {
                                    SBDSB = rowStatSBDSB[0].ToString();
                                }
                            }

                            using (DataTable dtstat = new DataTable())
                            {
                                string sql_search_stat = "SELECT STS FROM tblSBReceive WHERE EID='" + eid + "'";
                                using (OleDbDataAdapter adapter = new OleDbDataAdapter(sql_search_stat, conn))
                                {
                                    adapter.Fill(dtstat);
                                }
                                totalReceive = dtstat.Rows.Count;
                                foreach (DataRow rowStat in dtstat.Rows)
                                {
                                    if (rowStat[0].ToString() == "N")
                                    {
                                        dispatchFlag = true;
                                    }
                                    else { dispatchFlag = false; }
                                }

                                if (dispatchFlag && (totalDispatch > totalReceive))
                                {
                                    string sql_insrt = @"INSERT INTO tblSBReceive (DRDate, DRTime, EID, Remarks, SBDSB, STS)
                                                     VALUES ('" + date + "','" + time + "','" + eid + "','Receive','" + SBDSB + "','" + sts + "')";
                                    conn.Open();
                                    //MessageBox.Show(sql_insrt);
                                    using (OleDbCommand cmd = new OleDbCommand(sql_insrt, conn))
                                    {
                                        cmd.ExecuteNonQuery();
                                    }
                                    conn.Close();
                                    label29.ForeColor = Color.Maroon;
                                    label29.Text = "Receive Successful (NEGATIVE)\nEnrollment ID: " + eid + "\nDate: " + date + "\nTime: " + time + "\n" + SBDSB;
                                    conn.Close();
                                }
                                else
                                {
                                    label29.ForeColor = Color.DarkRed;
                                    label29.Text ="Already Received";
                                }
                            }
                        }

                    }

                }
                textBox5.Text = "";
            }
            else { MessageBox.Show("ENROLLMENT ID NOT GIVEN PROPERLY"); }
        }

        private void button10_Click_2(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        private void button9_Click_1(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            label3.Text = "EID: " + textBox4.Text;
            label11.Text = "EID: " + textBox4.Text;
            showQueryResult();            
            
            
            textBox4.Text = "";
            textBox4.Focus();
        }
        
        private void button8_Click_1(object sender, EventArgs e)
        {
           if (update_form == "dispatch")
            {
                Form update_dp = new update_dp(dtDSB, dataGridSelectedDSB, dataGridSelectedLootNumber, dataGridSelectedDID, this, dataGridSelectedEID, dataGridView1.Rows.Count > dataGridView2.Rows.Count, dataGridSelectedIsByHand);
                update_dp.Show();
                button8.Enabled = false;
            }
            else if (update_form == "receive")
            {
                Form update_ri = new updateRI(dtDSB, dataGridSelectedDSB, dataGridSelectedLootNumber, dataGridSelectedDID, this);
                update_ri.Show();
                button8.Enabled = false;
            }
        }

        private void tableLayoutPanel1_MouseClick(object sender, MouseEventArgs e)
        {
           
        }

        private void tableLayoutPanel5_MouseClick(object sender, MouseEventArgs e)
        {
           
        }

        private void tableLayoutPanel11_MouseClick(object sender, MouseEventArgs e)
        {
        }

        private void tableLayoutPanel6_MouseClick(object sender, MouseEventArgs e)
        {
            
        }

        private void tableLayoutPanel15_MouseClick(object sender, MouseEventArgs e)
        {
            
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && (e.ColumnIndex == 8 || e.ColumnIndex == 0))
            {
               
                DataGridViewRow row = this.dataGridView1.Rows[e.RowIndex];               
                dataGridSelectedLootNumber = row.Cells["LOT No"].Value.ToString();
                dataGridSelectedDID = row.Cells["DID"].Value.ToString();
                dataGridSelectedDSB = row.Cells["SBDSB"].Value.ToString();
                dataGridSelectedEID = row.Cells["EID"].Value.ToString();
                if(row.Cells["Dispatch"].Value.ToString().Trim() == "By Hand")
                {
                    dataGridSelectedIsByHand = "1";
                }
                else {dataGridSelectedIsByHand = "0";}
                
                /*foreach (DataGridViewRow r in dataGridView1.Rows)
                {
                    r.Cells[0].Style.ForeColor = Color.Black;
                    r.Cells[1].Style.ForeColor = Color.Black;
                    r.Cells[2].Style.ForeColor = Color.Black;
                    r.Cells[3].Style.ForeColor = Color.Black;
                    r.Cells[4].Style.ForeColor = Color.Black;
                }
                foreach (DataGridViewRow r in dataGridView2.Rows)
                {
                    r.Cells[0].Style.ForeColor = Color.Black;
                    r.Cells[1].Style.ForeColor = Color.Black;
                    r.Cells[2].Style.ForeColor = Color.Black;
                    //r.Cells[3].Style.ForeColor = Color.Black;
                }
                dataGridView1.Rows[e.RowIndex].Cells[0].Style.ForeColor = Color.Red;
                dataGridView1.Rows[e.RowIndex].Cells[1].Style.ForeColor = Color.Red;
                dataGridView1.Rows[e.RowIndex].Cells[2].Style.ForeColor = Color.Red;
                dataGridView1.Rows[e.RowIndex].Cells[3].Style.ForeColor = Color.Red;
                dataGridView1.Rows[e.RowIndex].Cells[4].Style.ForeColor = Color.Red;*/
                Form update_dp = new update_dp(dtDSB, dataGridSelectedDSB, dataGridSelectedLootNumber, dataGridSelectedDID, this, dataGridSelectedEID, dataGridView1.Rows.Count > dataGridView2.Rows.Count, dataGridSelectedIsByHand);
                update_dp.Show();
            }
        }

        private void dataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            //MessageBox.Show(e.ColumnIndex.ToString());
            if (e.RowIndex >= 0 && (e.ColumnIndex == 5 || e.ColumnIndex == 0))
            {
              
                DataGridViewRow row = this.dataGridView2.Rows[e.RowIndex];
                dataGridSelectedLootNumber = row.Cells["RESULT"].Value.ToString();
                dataGridSelectedDID = row.Cells["RID"].Value.ToString();
                dataGridSelectedDSB = row.Cells["SBDSB"].Value.ToString();                
                /*foreach (DataGridViewRow r in dataGridView2.Rows)
                {
                    r.Cells[0].Style.ForeColor = Color.Black;
                    r.Cells[1].Style.ForeColor = Color.Black;
                    r.Cells[2].Style.ForeColor = Color.Black;
                    //r.Cells[3].Style.ForeColor = Color.Black;
                }
                foreach (DataGridViewRow r in dataGridView1.Rows)
                {
                    r.Cells[0].Style.ForeColor = Color.Black;
                    r.Cells[1].Style.ForeColor = Color.Black;
                    r.Cells[2].Style.ForeColor = Color.Black;
                    r.Cells[3].Style.ForeColor = Color.Black;
                    r.Cells[4].Style.ForeColor = Color.Black;
                }
                dataGridView2.Rows[e.RowIndex].Cells[0].Style.ForeColor = Color.Red;
                dataGridView2.Rows[e.RowIndex].Cells[1].Style.ForeColor = Color.Red;
                dataGridView2.Rows[e.RowIndex].Cells[2].Style.ForeColor = Color.Red;
                dataGridView2.Rows[e.RowIndex].Cells[3].Style.ForeColor = Color.Red;*/

                Form update_ri = new updateRI(dtDSB, dataGridSelectedDSB, dataGridSelectedLootNumber, dataGridSelectedDID, this);
                update_ri.Show();
            }
        }
        public static bool update_reactive = false;
        private void Form1_Activated(object sender, EventArgs e)
        {
            if (textBox4.Text.Length==15)
            {
                showQueryResult();
            }
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            const char Delete = (char)8;
            e.Handled = !Char.IsDigit(e.KeyChar) && e.KeyChar != Delete;
        }

        private void textBox3_KeyPress(object sender, KeyPressEventArgs e)
        {
            const char Delete = (char)8;
            e.Handled = !Char.IsDigit(e.KeyChar) && e.KeyChar != Delete;
        }

        private void textBox5_KeyPress(object sender, KeyPressEventArgs e)
        {
            const char Delete = (char)8;
            e.Handled = !Char.IsDigit(e.KeyChar) && e.KeyChar != Delete;
        }

        private void textBox4_KeyPress(object sender, KeyPressEventArgs e)
        {
            const char Delete = (char)8;
            e.Handled = !Char.IsDigit(e.KeyChar) && e.KeyChar != Delete;
        }

        private void textBox1_KeyPress_1(object sender, KeyPressEventArgs e)
        {
            const char Delete = (char)8;
            e.Handled = !Char.IsDigit(e.KeyChar) && e.KeyChar != Delete;
            
        }

        private void tabControl1_Selected(object sender, TabControlEventArgs e)
        {
          
        }

        void deleteFile()
        {
            System.IO.DirectoryInfo di = new DirectoryInfo("temp");
            System.IO.Directory.CreateDirectory("temp");
            foreach (FileInfo file in di.GetFiles())
            {
                try
                {
                    file.Delete();
                }
                catch (Exception e)
                {
 
                }
            }
        }


        private void button11_Click_1(object sender, EventArgs e)
        {
            deleteFile();
            string startDate = dateTimePicker5.Value.ToString("MM-dd-yyyy");
            string endDate = dateTimePicker6.Value.ToString("MM-dd-yyyy");
            string condition = "";
            string SBDSB = comboBox1.Text;
            string loot = comboBox7.Text;
            string fileName = "";
            if (loot == "")
            {
                MessageBox.Show("Please Select Loot/Reference Number");
            }            
            else
            {
                if (loot == "No Loot/Reference Number")
                {
                    loot = "";
                }
                if (comboBox8.Text == "DISPATCH")
                {
                    condition = "STS='ND' AND SBDSB='" + SBDSB + "' AND ";
                    fileName = "dispatch_";
                }
                else if (comboBox8.Text == "REDISPATCH")
                {
                    condition = "STS='RD' AND SBDSB='" + SBDSB + "' AND ";
                    fileName = "redispatch_";
                }
                else
                {
                    condition = "SBDSB='" + SBDSB + "' AND ";
                    fileName = "all_dispatch_";
                }
                fileName += SBDSB + "_" + startDate + "_" + endDate;
                fileName = fileName.Replace('/', '.');
                SaveFileDialog saveFileDialog1 = new SaveFileDialog();
                saveFileDialog1.Filter = "PDF(*.pdf)|*.pdf";
                saveFileDialog1.FilterIndex = 2;
                saveFileDialog1.RestoreDirectory = true;
               
                saveFileDialog1.FileName = fileName;
                string savePath = @"temp\" + fileName + "_" + DateTime.Now.Ticks.ToString() + ".pdf";
                try
                {
                    //if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                    {                        
                        Document doc = new Document(iTextSharp.text.PageSize.A4, 10, 10, 42, 35);
                        PdfWriter wri = PdfWriter.GetInstance(doc, new FileStream(savePath, FileMode.Create));
                        doc.Open();

                        PdfPTable table = new PdfPTable(5);
                        iTextSharp.text.Font fontH1 = new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.TIMES_ROMAN, 16, iTextSharp.text.Font.BOLD);
                        iTextSharp.text.Font fontH2 = new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.TIMES_ROMAN, 12, iTextSharp.text.Font.BOLD);
                        iTextSharp.text.Font fontH3 = new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.TIMES_ROMAN, 10, iTextSharp.text.Font.NORMAL);
                        iTextSharp.text.Font fontH4 = new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.TIMES_ROMAN, 12, iTextSharp.text.Font.UNDERLINE);
                        iTextSharp.text.Font fontH5 = new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.TIMES_ROMAN, 11, iTextSharp.text.Font.NORMAL);

                        var cell = new PdfPCell(new Phrase("Government of the People's Republic of Bangladesh", fontH1));
                        cell.Colspan = 5;
                        cell.HorizontalAlignment = 1;
                        cell.VerticalAlignment = 1;
                        cell.BorderColor = BaseColor.WHITE;
                        table.AddCell(cell);

                        cell = new PdfPCell(new Phrase(Form1.sb_name_static, fontH2));
                        cell.Colspan = 5;
                        cell.HorizontalAlignment = 1;
                        cell.VerticalAlignment = 1;
                        cell.BorderColor = BaseColor.WHITE;
                        table.AddCell(cell);

                        cell = new PdfPCell(new Phrase("Dispatch Report", fontH2));
                        cell.Colspan = 5;
                        cell.HorizontalAlignment = 1;
                        cell.VerticalAlignment = 1;
                        cell.BorderColor = BaseColor.WHITE;
                        table.AddCell(cell);



                        cell = new PdfPCell(new Phrase("Print Date : " + DateTime.Today.ToString("d MMM yyyy"), fontH3));
                        cell.Colspan = 5;
                        cell.HorizontalAlignment = 2;
                        cell.VerticalAlignment = 1;
                        cell.BorderColor = BaseColor.WHITE;
                        table.AddCell(cell);

                        cell = new PdfPCell(new Phrase("SB/DSB : " + SBDSB, fontH3));
                        cell.HorizontalAlignment = 0;
                        cell.VerticalAlignment = 1;
                        cell.BorderColor = BaseColor.WHITE;
                        table.AddCell(cell);

                        cell = new PdfPCell(new Phrase("From : " + dateTimePicker5.Value.ToString("d MMM yyyy") + "    To : " + dateTimePicker6.Value.ToString("d MMM yyyy"), fontH3));
                        cell.Colspan = 4;
                        cell.HorizontalAlignment = 1;
                        cell.VerticalAlignment = 1;
                        cell.BorderColor = BaseColor.WHITE;
                        table.AddCell(cell);

                        cell = new PdfPCell(new Phrase("Lot/Ref No: " + comboBox7.Text, fontH3));
                        cell.Colspan = 4;
                        cell.HorizontalAlignment = 0;
                        cell.VerticalAlignment = 1;
                        cell.BorderColor = BaseColor.WHITE;
                        table.AddCell(cell);



                        using (DataTable dt = new DataTable())
                        {

                            string sql_search = "SELECT EID,DRDate,DRTime,SBDSB,LotNo,STS FROM tblSBDispatch WHERE " + condition + " DRDate BETWEEN #" + startDate + "# AND #" + endDate + "# AND LotNo='" + loot.Trim() + "' AND isByHand=0  ORDER BY STS";
                            //MessageBox.Show(sql_search);
                            using (OleDbDataAdapter adapter = new OleDbDataAdapter(sql_search, conn))
                            {
                                adapter.Fill(dt);
                            }

                            if (comboBox8.Text != "REDISPATCH")
                            {
                                cell = new PdfPCell(new Phrase("Total : " + dt.Rows.Count, fontH2));
                                cell.HorizontalAlignment = 1;
                                cell.VerticalAlignment = 1;
                                cell.BorderColor = BaseColor.WHITE;
                                table.AddCell(cell);


                                cell = new PdfPCell(new Phrase("New Dispatch", fontH4));
                                cell.Colspan = 5;
                                cell.HorizontalAlignment = 0;
                                cell.VerticalAlignment = 1;
                                cell.BorderColor = BaseColor.WHITE;
                                table.AddCell(cell);

                                cell = new PdfPCell(new Phrase(" ", fontH2));
                                cell.Colspan = 5;
                                cell.HorizontalAlignment = 1;
                                cell.VerticalAlignment = 1;
                                cell.BorderColor = BaseColor.WHITE;
                                table.AddCell(cell);

                                table.AddCell("Enrollment ID");
                                table.AddCell("Enrollment ID");
                                table.AddCell("Enrollment ID");
                                table.AddCell("Enrollment ID");
                                table.AddCell("Enrollment ID");
                            }

                            Boolean RDFlage = true;
                            int count = 0;
                            foreach (DataRow row in dt.Rows)
                            {
                                if (RDFlage && row[5].ToString() == "RD")
                                {
                                    int blankcellsize = count % 5;
                                    if (blankcellsize > 0)
                                    {
                                        for (int i = 0; i < 5 - blankcellsize; i++)
                                        {
                                            cell = new PdfPCell(new Phrase(" "));
                                            cell.HorizontalAlignment = 1;
                                            cell.VerticalAlignment = 1;
                                            cell.BorderColor = BaseColor.BLACK;
                                            table.AddCell(cell);
                                        }
                                    }
                                    if (comboBox8.Text != "REDISPATCH")
                                    {
                                        cell = new PdfPCell(new Phrase("Total : " + count, fontH2));
                                        cell.Colspan = 5;
                                        cell.HorizontalAlignment = 2;
                                        cell.VerticalAlignment = 1;
                                        cell.BorderColor = BaseColor.WHITE;
                                        table.AddCell(cell);
                                    }
                                    else
                                    {
                                        cell = new PdfPCell(new Phrase("Total : " + dt.Rows.Count, fontH2));
                                        cell.Colspan = 5;
                                        cell.HorizontalAlignment = 2;
                                        cell.VerticalAlignment = 1;
                                        cell.BorderColor = BaseColor.WHITE;
                                        table.AddCell(cell);
                                    }
                                    cell = new PdfPCell(new Phrase("Redispatch", fontH4));
                                    cell.Colspan = 5;
                                    cell.HorizontalAlignment = 0;
                                    cell.VerticalAlignment = 1;
                                    cell.BorderColor = BaseColor.WHITE;
                                    table.AddCell(cell);

                                    cell = new PdfPCell(new Phrase(" ", fontH2));
                                    cell.Colspan = 5;
                                    cell.HorizontalAlignment = 1;
                                    cell.VerticalAlignment = 1;
                                    cell.BorderColor = BaseColor.WHITE;
                                    table.AddCell(cell);

                                    table.AddCell("Enrollment ID");
                                    table.AddCell("Enrollment ID");
                                    table.AddCell("Enrollment ID");
                                    table.AddCell("Enrollment ID");
                                    table.AddCell("Enrollment ID");
                                    RDFlage = false;
                                    count = 0;
                                }

                                cell = new PdfPCell(new Phrase(row[0].ToString(), fontH5));
                                cell.HorizontalAlignment = 1;
                                cell.VerticalAlignment = 1;
                                cell.BorderColor = BaseColor.BLACK;
                                table.AddCell(cell);
                                count++;
                            }

                            int blankcellsize1 = count % 5;
                            if (blankcellsize1 > 0)
                            {
                                for (int i = 0; i < 5 - blankcellsize1; i++)
                                {
                                    cell = new PdfPCell(new Phrase(" "));
                                    cell.HorizontalAlignment = 1;
                                    cell.VerticalAlignment = 1;
                                    cell.BorderColor = BaseColor.BLACK;
                                    table.AddCell(cell);
                                }
                            }

                            cell = new PdfPCell(new Phrase("Total : " + count, fontH2));
                            cell.Colspan = 5;
                            cell.HorizontalAlignment = 2;
                            cell.VerticalAlignment = 1;
                            cell.Border =  PdfPCell.TOP_BORDER;
                            cell.BorderColor = BaseColor.BLACK;
                            table.AddCell(cell);


                            //////
                            cell = new PdfPCell(new Phrase(" ", fontH2));
                            cell.Colspan = 5;
                            cell.HorizontalAlignment = 1;
                            cell.VerticalAlignment = 1;
                            cell.BorderColor = BaseColor.WHITE;
                            table.AddCell(cell);

                            cell = new PdfPCell(new Phrase(" ", fontH2));
                            cell.Colspan = 5;
                            cell.HorizontalAlignment = 1;
                            cell.VerticalAlignment = 1;
                            cell.BorderColor = BaseColor.WHITE;
                            table.AddCell(cell);

                            cell = new PdfPCell(new Phrase(" ", fontH2));
                            cell.Colspan = 5;
                            cell.HorizontalAlignment = 1;
                            cell.VerticalAlignment = 1;
                            cell.BorderColor = BaseColor.WHITE;
                            table.AddCell(cell);

                            cell = new PdfPCell(new Phrase(" ", fontH2));
                            cell.Colspan = 5;
                            cell.HorizontalAlignment = 1;
                            cell.VerticalAlignment = 1;
                            cell.BorderColor = BaseColor.WHITE;
                            table.AddCell(cell);

                            cell = new PdfPCell(new Phrase(" ", fontH2));
                            cell.Colspan = 5;
                            cell.HorizontalAlignment = 1;
                            cell.VerticalAlignment = 1;
                            cell.BorderColor = BaseColor.WHITE;
                            table.AddCell(cell);

                            cell = new PdfPCell(new Phrase(" ", fontH2));
                            cell.Colspan = 5;
                            cell.HorizontalAlignment = 1;
                            cell.VerticalAlignment = 1;
                            cell.BorderColor = BaseColor.WHITE;
                            table.AddCell(cell);

                            cell = new PdfPCell(new Phrase(" ", fontH2));
                            cell.Colspan = 5;
                            cell.HorizontalAlignment = 1;
                            cell.VerticalAlignment = 1;
                            cell.BorderColor = BaseColor.WHITE;
                            table.AddCell(cell);
                            //////

                            cell = new PdfPCell(new Phrase("Sender's Signature", fontH2));
                            cell.Colspan = 2;
                            cell.HorizontalAlignment = 1;
                            cell.VerticalAlignment = 1;
                            cell.BorderColor = BaseColor.WHITE;
                            table.AddCell(cell);

                            cell = new PdfPCell(new Phrase(" ", fontH2));
                            cell.HorizontalAlignment = 1;
                            cell.VerticalAlignment = 1;
                            cell.BorderColor = BaseColor.WHITE;
                            table.AddCell(cell);

                            cell = new PdfPCell(new Phrase("Receiver's Signature", fontH2));
                            cell.Colspan = 2;
                            cell.HorizontalAlignment = 1;
                            cell.VerticalAlignment = 1;
                            cell.BorderColor = BaseColor.WHITE;
                            table.AddCell(cell);

                        }
                        doc.Add(table);
                        doc.Close();
                        System.Diagnostics.Process.Start(savePath);
                    }
                }
                catch (Exception ex) { MessageBox.Show(ex.Message); }
            }
        }

        private void button12_Click(object sender, EventArgs e)
        {
            string startDate = dateTimePicker8.Value.ToString("d");
            string endDate = dateTimePicker7.Value.ToString("d");
            string condition = "";
            string SBDSB = comboBox11.Text;
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            Stream myStream;
            saveFileDialog1.Filter = "PDF(*.pdf)|*.pdf";
            saveFileDialog1.FilterIndex = 2;
            saveFileDialog1.RestoreDirectory = true;
            string fileName = SBDSB + "_";
            if (comboBox9.Text == "POSITIVE")
            {
                condition = "STS='P' AND SBDSB='" + SBDSB + "' AND";
                fileName += "receive_positive_";
            }
            else if (comboBox9.Text == "NEGATIVE")
            {
                condition = "STS='N' AND SBDSB='" + SBDSB + "' AND";
                fileName += "receive_negative_";
            }
            else
            {
                condition = "SBDSB='" + SBDSB + "' AND";
                fileName += "receive_all_";
            }
            fileName += "from_" + startDate.Replace('/', '.') + "_to_" + endDate.Replace('/', '.');
            string savePath = @"temp\" + fileName + "_" + DateTime.Now.Ticks.ToString() + ".pdf";
            try
            {
                //if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    
                    Document doc = new Document(iTextSharp.text.PageSize.A4, 10, 10, 42, 35);
                    PdfWriter wri = PdfWriter.GetInstance(doc, new FileStream(savePath, FileMode.Create));
                    doc.Open();

                    PdfPTable table = new PdfPTable(5);
                    iTextSharp.text.Font fontH1 = new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.TIMES_ROMAN, 16, iTextSharp.text.Font.BOLD);
                    iTextSharp.text.Font fontH2 = new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.TIMES_ROMAN, 12, iTextSharp.text.Font.BOLD);
                    iTextSharp.text.Font fontH3 = new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.TIMES_ROMAN, 10, iTextSharp.text.Font.NORMAL);
                    iTextSharp.text.Font fontH4 = new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.TIMES_ROMAN, 12, iTextSharp.text.Font.UNDERLINE);
                    iTextSharp.text.Font fontH5 = new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.TIMES_ROMAN, 11, iTextSharp.text.Font.NORMAL);

                    var cell = new PdfPCell(new Phrase("Government of the People's Republic of Bangladesh", fontH1));
                    cell.Colspan = 5;
                    cell.HorizontalAlignment = 1;
                    cell.VerticalAlignment = 1;
                    cell.BorderColor = BaseColor.WHITE;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(Form1.sb_name_static, fontH2));
                    cell.Colspan = 5;
                    cell.HorizontalAlignment = 1;
                    cell.VerticalAlignment = 1;
                    cell.BorderColor = BaseColor.WHITE;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase("Receive Report", fontH2));
                    cell.Colspan = 5;
                    cell.HorizontalAlignment = 1;
                    cell.VerticalAlignment = 1;
                    cell.BorderColor = BaseColor.WHITE;
                    table.AddCell(cell);



                    cell = new PdfPCell(new Phrase("Print Date : " + DateTime.Today.ToString("d MMM yyyy"), fontH3));
                    cell.Colspan = 5;
                    cell.HorizontalAlignment = 2;
                    cell.VerticalAlignment = 1;
                    cell.BorderColor = BaseColor.WHITE;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase("SB/DSB : " + SBDSB, fontH3));
                    cell.HorizontalAlignment = 0;
                    cell.VerticalAlignment = 1;
                    cell.BorderColor = BaseColor.WHITE;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase("From : " + dateTimePicker8.Value.ToString("d MMM yyyy") + "    To : " + dateTimePicker7.Value.ToString("d MMM yyyy"), fontH3));
                    cell.Colspan = 3;
                    cell.HorizontalAlignment = 1;
                    cell.VerticalAlignment = 1;
                    cell.BorderColor = BaseColor.WHITE;
                    table.AddCell(cell);


                    int count = 0;
                    using (DataTable dt = new DataTable())
                    {

                        string sql_search = "SELECT EID,DRDate,DRTime,SBDSB,STS FROM tblSBReceive WHERE " + condition + " DRDate BETWEEN #" + startDate + "# AND #" + endDate + "# ORDER BY STS DESC";
                        //MessageBox.Show(sql_search);
                        using (OleDbDataAdapter adapter = new OleDbDataAdapter(sql_search, conn))
                        {
                            try
                            {
                                adapter.Fill(dt);
                            }
                            catch (Exception ex) { }
                        }

                        if (comboBox9.Text != "NEGATIVE")
                        {
                            cell = new PdfPCell(new Phrase("Total : " + dt.Rows.Count, fontH2));
                            cell.HorizontalAlignment = 1;
                            cell.VerticalAlignment = 1;
                            cell.BorderColor = BaseColor.WHITE;
                            table.AddCell(cell);
                        
                            cell = new PdfPCell(new Phrase("Positive", fontH4));
                            cell.Colspan = 5;
                            cell.HorizontalAlignment = 0;
                            cell.VerticalAlignment = 1;
                            cell.BorderColor = BaseColor.WHITE;
                            table.AddCell(cell);

                            cell = new PdfPCell(new Phrase(" ", fontH2));
                            cell.Colspan = 5;
                            cell.HorizontalAlignment = 1;
                            cell.VerticalAlignment = 1;
                            cell.BorderColor = BaseColor.WHITE;
                            table.AddCell(cell);

                            table.AddCell("Enrollment ID");
                            table.AddCell("Enrollment ID");
                            table.AddCell("Enrollment ID");
                            table.AddCell("Enrollment ID");
                            table.AddCell("Enrollment ID");
                        }
                        Boolean PNFlage = true; 
                        foreach (DataRow row in dt.Rows)
                        {
                            if (PNFlage && row[4].ToString() == "N")
                            {
                                int blankcellsize = count % 5;
                                if (blankcellsize > 0)
                                {
                                    for (int i = 0; i < 5 - blankcellsize; i++)
                                    {
                                        cell = new PdfPCell(new Phrase(" "));
                                        cell.HorizontalAlignment = 1;
                                        cell.VerticalAlignment = 1;
                                        table.AddCell(cell);
                                    }
                                }
                                if (comboBox9.Text != "NEGATIVE")
                                {
                                    cell = new PdfPCell(new Phrase("Total : " + count, fontH2));
                                    cell.Colspan = 5;
                                    cell.HorizontalAlignment = 2;
                                    cell.VerticalAlignment = 1;
                                    cell.BorderColor = BaseColor.WHITE;
                                    table.AddCell(cell);
                                }
                                else 
                                {
                                    cell = new PdfPCell(new Phrase("Total : " + dt.Rows.Count, fontH2));
                                    cell.Colspan = 5;
                                    cell.HorizontalAlignment = 2;
                                    cell.VerticalAlignment = 1;
                                    cell.BorderColor = BaseColor.WHITE;
                                    table.AddCell(cell);
                                }
                                cell = new PdfPCell(new Phrase("Negative", fontH4));
                                cell.Colspan = 5;
                                cell.HorizontalAlignment = 0;
                                cell.VerticalAlignment = 1;
                                cell.BorderColor = BaseColor.WHITE;
                                table.AddCell(cell);

                                cell = new PdfPCell(new Phrase(" ", fontH2));
                                cell.Colspan = 5;
                                cell.HorizontalAlignment = 1;
                                cell.VerticalAlignment = 1;
                                cell.BorderColor = BaseColor.WHITE;
                                table.AddCell(cell);

                                table.AddCell("Enrollment ID");
                                table.AddCell("Enrollment ID");
                                table.AddCell("Enrollment ID");
                                table.AddCell("Enrollment ID");
                                table.AddCell("Enrollment ID");

                                PNFlage = false;
                                count = 0;
                            }

                            cell = new PdfPCell(new Phrase(row[0].ToString(), fontH5));
                            cell.HorizontalAlignment = 1;
                            cell.VerticalAlignment = 1;
                            cell.BorderColor = BaseColor.BLACK;                    
                            table.AddCell(cell);
                            count++;
                        }

                        int blankcellsize1 = count % 5;
                        if (blankcellsize1 > 0)
                        {
                            for (int i = 0; i < 5 - blankcellsize1; i++)
                            {
                                cell = new PdfPCell(new Phrase(" "));
                                cell.HorizontalAlignment = 1;
                                cell.VerticalAlignment = 1;
                                table.AddCell(cell);
                            }
                        }

                        cell = new PdfPCell(new Phrase("Total : " + count, fontH2));
                        cell.Colspan = 5;
                        cell.HorizontalAlignment = 2;
                        cell.VerticalAlignment = 1;
                        cell.Border = PdfPCell.TOP_BORDER;
                        cell.BorderColor = BaseColor.BLACK;
                        table.AddCell(cell);


                        //////
                        cell = new PdfPCell(new Phrase(" ", fontH2));
                        cell.Colspan = 5;
                        cell.HorizontalAlignment = 1;
                        cell.VerticalAlignment = 1;
                        cell.BorderColor = BaseColor.WHITE;
                        table.AddCell(cell);

                        cell = new PdfPCell(new Phrase(" ", fontH2));
                        cell.Colspan = 5;
                        cell.HorizontalAlignment = 1;
                        cell.VerticalAlignment = 1;
                        cell.BorderColor = BaseColor.WHITE;
                        table.AddCell(cell);

                        cell = new PdfPCell(new Phrase(" ", fontH2));
                        cell.Colspan = 5;
                        cell.HorizontalAlignment = 1;
                        cell.VerticalAlignment = 1;
                        cell.BorderColor = BaseColor.WHITE;
                        table.AddCell(cell);

                        cell = new PdfPCell(new Phrase(" ", fontH2));
                        cell.Colspan = 5;
                        cell.HorizontalAlignment = 1;
                        cell.VerticalAlignment = 1;
                        cell.BorderColor = BaseColor.WHITE;
                        table.AddCell(cell);

                        cell = new PdfPCell(new Phrase(" ", fontH2));
                        cell.Colspan = 5;
                        cell.HorizontalAlignment = 1;
                        cell.VerticalAlignment = 1;
                        cell.BorderColor = BaseColor.WHITE;
                        table.AddCell(cell);

                        cell = new PdfPCell(new Phrase(" ", fontH2));
                        cell.Colspan = 5;
                        cell.HorizontalAlignment = 1;
                        cell.VerticalAlignment = 1;
                        cell.BorderColor = BaseColor.WHITE;
                        table.AddCell(cell);

                        cell = new PdfPCell(new Phrase(" ", fontH2));
                        cell.Colspan = 5;
                        cell.HorizontalAlignment = 1;
                        cell.VerticalAlignment = 1;
                        cell.BorderColor = BaseColor.WHITE;
                        table.AddCell(cell);
                        //////

                        cell = new PdfPCell(new Phrase("Sender's Signature", fontH2));
                        cell.Colspan = 2;
                        cell.HorizontalAlignment = 1;
                        cell.VerticalAlignment = 1;
                        cell.BorderColor = BaseColor.WHITE;
                        table.AddCell(cell);

                        cell = new PdfPCell(new Phrase(" ", fontH2));
                        cell.HorizontalAlignment = 1;
                        cell.VerticalAlignment = 1;
                        cell.BorderColor = BaseColor.WHITE;
                        table.AddCell(cell);

                        cell = new PdfPCell(new Phrase("Receiver's Signature", fontH2));
                        cell.Colspan = 2;
                        cell.HorizontalAlignment = 1;
                        cell.VerticalAlignment = 1;
                        cell.BorderColor = BaseColor.WHITE;
                        table.AddCell(cell);
                    }
                    doc.Add(table);
                    doc.Close();
                    System.Diagnostics.Process.Start(savePath);
                }
            
            }
            catch (Exception ex) { MessageBox.Show("File is Already Open"); }
        }

    private void tabControl1_DrawItem(object sender, DrawItemEventArgs e)
    {
        //get tabpage
        TabPage tabPages = tabControl1.TabPages[e.Index];
        Graphics graphics = e.Graphics;
        Brush textBrush = new SolidBrush(Color.White); //fore color brush
        System.Drawing.Rectangle tabBounds = tabControl1.GetTabRect(e.Index);
        if (e.State == DrawItemState.Selected)
        {
            graphics.FillRectangle(Brushes.DarkGray, e.Bounds); //fill background color
        }
        else
        {
            textBrush = new System.Drawing.SolidBrush(e.ForeColor);
            e.DrawBackground();
        }
        System.Drawing.Font tabFont = new System.Drawing.Font("Agency FB", 20, FontStyle.Regular | FontStyle.Regular, GraphicsUnit.Pixel);
        StringFormat strFormat = new StringFormat();
        strFormat.Alignment = StringAlignment.Center;
        strFormat.LineAlignment = StringAlignment.Center;

        graphics.DrawString(tabPages.Text, tabFont, textBrush, tabBounds, new StringFormat(strFormat));
        graphics.Dispose();
        textBrush.Dispose();
       
    }

    private void tableLayoutPanel3_Paint(object sender, PaintEventArgs e)
    {

    }

    private void tableLayoutPanel3_CellPaint(object sender, TableLayoutCellPaintEventArgs e)
    {
        var rectangle = e.CellBounds;
        rectangle.Inflate(-1, -1);

        ControlPaint.DrawBorder3D(e.Graphics, rectangle, Border3DStyle.Raised, Border3DSide.Right); // 3D border
        //ControlPaint.DrawBorder(e.Graphics, rectangle, Color.Red, ButtonBorderStyle.Dotted); // dotted border
    }

    private void tableLayoutPanel2_CellPaint(object sender, TableLayoutCellPaintEventArgs e)
    {
        var rectangle = e.CellBounds;
        rectangle.Inflate(-1, -1);

        ControlPaint.DrawBorder3D(e.Graphics, rectangle, Border3DStyle.Raised, Border3DSide.Right); // 3D border
    }

    private void tableLayoutPanel10_CellPaint(object sender, TableLayoutCellPaintEventArgs e)
    {
        var rectangle = e.CellBounds;
        rectangle.Inflate(-1, -1);

        ControlPaint.DrawBorder3D(e.Graphics, rectangle, Border3DStyle.Raised, Border3DSide.Right); // 3D border
    }

    private void button9_DragEnter(object sender, DragEventArgs e)
    {
        
    }

    private void button9_Enter(object sender, EventArgs e)
    {
       
    }

    private void button9_Leave(object sender, EventArgs e)
    {
        
    }

    private void button9_MouseEnter(object sender, EventArgs e)
    {
        button9.BackColor = Color.Maroon;
        button9.ForeColor = Color.White;
        button9.FlatAppearance.BorderColor = Color.Maroon;
    }

    private void button9_MouseLeave(object sender, EventArgs e)
    {
        button9.BackColor = Color.Transparent;
        button9.ForeColor = Color.Black;
    }

    private void button9_MouseEnter_1(object sender, EventArgs e)
    {
        button9.BackColor = Color.Maroon;
        button9.ForeColor = Color.White;
        button9.FlatAppearance.BorderColor = Color.Maroon;
    }

    private void button9_MouseLeave_1(object sender, EventArgs e)
    {
        button9.BackColor = Color.Transparent;
        button9.ForeColor = Color.Maroon;
    }

    private void button9_Click_2(object sender, EventArgs e)
    {
        Application.Exit();
    }

    private void button10_Click_3(object sender, EventArgs e)
    {
        this.WindowState = FormWindowState.Minimized;
    }

    private void button1_MouseEnter(object sender, EventArgs e)
    {
        button1.BackColor = Color.FromArgb(0, 50, 51);
        button1.ForeColor = Color.Transparent;
        button1.FlatAppearance.BorderColor = Color.FromArgb(0, 50, 51);
    }

    private void button1_MouseLeave(object sender, EventArgs e)
    {
        button1.BackColor = Color.Transparent;
        button1.ForeColor = Color.FromArgb(0, 50, 51);
    }

    private void button2_MouseEnter(object sender, EventArgs e)
    {
        button2.BackColor = Color.DarkGreen;
        button2.ForeColor = Color.Transparent;
        button2.FlatAppearance.BorderColor = Color.DarkGreen;
    }

    private void button2_MouseLeave(object sender, EventArgs e)
    {
        button2.BackColor = Color.Transparent;
        button2.ForeColor = Color.DarkGreen;
    }

    private void button7_MouseEnter(object sender, EventArgs e)
    {
        button7.BackColor = Color.Maroon;
        button7.ForeColor = Color.Transparent;
        button7.FlatAppearance.BorderColor = Color.Maroon;
    }

    private void button7_MouseLeave(object sender, EventArgs e)
    {
        button7.BackColor = Color.Transparent;
        button7.ForeColor = Color.Maroon;
    }

    private void button3_MouseEnter(object sender, EventArgs e)
    {
        button3.BackColor = Color.FromArgb(0, 50, 51);
        button3.ForeColor = Color.Transparent;
        button3.FlatAppearance.BorderColor = Color.FromArgb(0, 50, 51);
    }

    private void button3_MouseLeave(object sender, EventArgs e)
    {
        button3.BackColor = Color.Transparent;
        button3.ForeColor = Color.FromArgb(0, 50, 51);
    }

    private void button11_MouseEnter(object sender, EventArgs e)
    {
        button11.BackColor = Color.FromArgb(0, 50, 51);
        button11.ForeColor = Color.Transparent;
        button11.FlatAppearance.BorderColor = Color.FromArgb(0, 50, 51);
    }

    private void button11_MouseLeave(object sender, EventArgs e)
    {
        button11.BackColor = Color.Transparent;
        button11.ForeColor = Color.FromArgb(0, 50, 51);
    }

    private void button12_MouseEnter(object sender, EventArgs e)
    {
        button12.BackColor = Color.FromArgb(0, 50, 51);
        button12.ForeColor = Color.Transparent;
        button12.FlatAppearance.BorderColor = Color.FromArgb(0, 50, 51);
    }

    private void button12_MouseLeave(object sender, EventArgs e)
    {
        button12.BackColor = Color.Transparent;
        button12.ForeColor = Color.FromArgb(0, 50, 51);
    }

    private void button4_MouseEnter(object sender, EventArgs e)
    {
        button4.BackColor = Color.FromArgb(0, 50, 51);
        button4.ForeColor = Color.Transparent;
        button4.FlatAppearance.BorderColor = Color.FromArgb(0, 50, 51);
    }

    private void button4_MouseLeave(object sender, EventArgs e)
    {
        button4.BackColor = Color.Transparent;
        button4.ForeColor = Color.FromArgb(0, 50, 51);
    }

    private void textBox1_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.KeyCode == Keys.Enter)
        {
            button1_Click_1(sender, e);
        }
    }

    private void textBox3_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.KeyCode == Keys.Enter)
        {
            button2_Click_1(sender, e);
        }
    }

    private void textBox5_KeyDown(object sender, KeyEventArgs e)
    {
         if (e.KeyCode == Keys.Enter)
        {
            button7_Click_1(sender, e);
        }        
    }

    private void textBox4_KeyDown(object sender, KeyEventArgs e)
    {
         if (e.KeyCode == Keys.Enter)
        {
            button3_Click_1(sender, e);
        }
    }

    private void textBox1_Enter(object sender, EventArgs e)
    {
        textBox1.Text = "";
    }

    private void textBox3_Enter(object sender, EventArgs e)
    {
        textBox3.Text = "";
    }

    private void textBox5_Enter(object sender, EventArgs e)
    {
        textBox5.Text = "";
    }

    private void textBox4_Enter(object sender, EventArgs e)
    {
        textBox4.Text = "";
    }

    private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
    {
        if (e.KeyChar == '\'')
        {
            e.Handled = true;
        }
    }

    private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
    {
        if (tabControl1.SelectedIndex == 0)
        {
            textBox1.Focus();
        }
        else if (tabControl1.SelectedIndex == 1)
        {
            textBox3.Focus();
        }
        else if (tabControl1.SelectedIndex == 2)
        {
            textBox5.Focus();
        }
        else if (tabControl1.SelectedIndex == 3)
        {
            textBox4.Focus();
        }        
    }

    private void dateTimePicker6_ValueChanged(object sender, EventArgs e)
    {
        string startDate = dateTimePicker5.Value.ToString("d");
        string endDate = dateTimePicker6.Value.ToString("d");
        string sb = comboBox1.Text;

        using (DataTable dt = new DataTable())
        {

            String sql = "SELECT DISTINCT LotNo FROM tblSBDispatch WHERE SBDSB='" + sb + "' AND DRDate BETWEEN #" + startDate + "# AND #" + endDate + "#";
            using (OleDbDataAdapter adapter = new OleDbDataAdapter(sql, conn))
            {
                adapter.Fill(dt);
            }
            comboBox7.Items.Clear();
            foreach (DataRow row in dt.Rows)
            {
                ComboboxItem item = new ComboboxItem();
                item.Text = row[0].ToString();
                item.Value = row[0].ToString();
                comboBox7.Items.Add(item);
            }
            ComboboxItem item1 = new ComboboxItem();
            item1.Text = "No Loot/Reference Number";
            item1.Value = "No Loot/Reference Number";
            comboBox7.Items.Add(item1);
        }
    }

    private void dateTimePicker5_ValueChanged(object sender, EventArgs e)
    {
        
        string startDate = dateTimePicker5.Value.ToString("d");
        string endDate = dateTimePicker6.Value.ToString("d");
        string sb = comboBox1.Text;

        using (DataTable dt = new DataTable())
        {

            String sql = "SELECT DISTINCT LotNo FROM tblSBDispatch WHERE SBDSB='" + sb + "' AND DRDate BETWEEN #" + startDate + "# AND #" + endDate + "#";
            using (OleDbDataAdapter adapter = new OleDbDataAdapter(sql, conn))
            {
                adapter.Fill(dt);
            }
            comboBox7.Items.Clear();
            foreach (DataRow row in dt.Rows)
            {
                ComboboxItem item = new ComboboxItem();
                item.Text = row[0].ToString();
                item.Value = row[0].ToString();
                comboBox7.Items.Add(item);
            }
            ComboboxItem item1 = new ComboboxItem();
            item1.Text = "No Loot/Reference Number";
            item1.Value = "No Loot/Reference Number";
            comboBox7.Items.Add(item1);
        }
    }

    private void dataGridView2_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
    {
        
    }

    private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
    {
        string startDate = dateTimePicker5.Value.ToString("d");
        string endDate = dateTimePicker6.Value.ToString("d");
        string sb = comboBox1.Text;

        using (DataTable dt = new DataTable())
        {

            String sql = "SELECT DISTINCT LotNo FROM tblSBDispatch WHERE SBDSB='" + sb + "' AND DRDate BETWEEN #" + startDate + "# AND #" + endDate + "#";
            using (OleDbDataAdapter adapter = new OleDbDataAdapter(sql, conn))
            {
                adapter.Fill(dt);
            }
            comboBox7.Items.Clear();
            foreach (DataRow row in dt.Rows)
            {
                ComboboxItem item = new ComboboxItem();
                item.Text = row[0].ToString();
                item.Value = row[0].ToString();
                comboBox7.Items.Add(item);
            }
            ComboboxItem item1 = new ComboboxItem();
            item1.Text = "No Loot/Reference Number";
            item1.Value = "No Loot/Reference Number";
            comboBox7.Items.Add(item1);
        }
    }

    private void button5_Click_1(object sender, EventArgs e)
    {
        String oldPass = textBox6.Text;
        String newPass = textBox7.Text;

        if (oldPass.Equals("") || newPass.Equals(""))
        {
            MessageBox.Show("Please fill up old and new password");
            return;
        }
        if (!oldPass.Equals(usercontrol.password))
        {
            MessageBox.Show("Old Password do not Match");
            return;
        }

        string sql_update = "UPDATE usercontrol SET `password`='" + newPass + "' WHERE `username`='" + usercontrol.username + "'";
        //MessageBox.Show(sql_update);
        conn.Open();
        OleDbCommand cmdUpdate = new OleDbCommand();
        cmdUpdate.Parameters.Clear();
        cmdUpdate.CommandText = sql_update;
        cmdUpdate.CommandType = CommandType.Text;
        cmdUpdate.Parameters.Add("@LIKeys", OleDbType.LongVarChar);
        cmdUpdate.Connection = conn;
        cmdUpdate.ExecuteNonQuery();
        conn.Close();
        usercontrol.password = newPass;
        MessageBox.Show("Pasword Change Successful");
        textBox6.Text="";
        textBox7.Text="";

    }

        
    }

    public class ComboboxItem
    {
        public string Text { get; set; }
        public object Value { get; set; }

        public override string ToString()
        {
            return Text;
        }
    }




}

