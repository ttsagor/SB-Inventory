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
    public partial class update_dp : Form
    {
        string DSB = "";
        string loot = "";
        string eid = "";
        Form f1;
        public update_dp(DataTable dt, string par_selected_index, string par_loot, string par_eid, Form form1, string dataGridSelectedEID, bool eidFlage, string dataGridSelectedIsByHand)
        {
            InitializeComponent();
            foreach (DataRow row in dt.Rows)
            {
                ComboboxItem item = new ComboboxItem();
                item.Text = row[0].ToString();
                item.Value = row[0].ToString();
                comboBox6.Items.Add(item);
            }

            textBox1.Text = dataGridSelectedEID;
            if (eidFlage)
            {
                textBox1.Enabled = true;
            }
            else { textBox1.Enabled = false; }
            DSB = par_selected_index;
            loot = par_loot;
            eid = par_eid;      
            textBox6.Text = loot;
            f1 = form1;
            comboBox6.SelectedIndex = comboBox6.FindString(par_selected_index);
            if (dataGridSelectedIsByHand == "1")
            {
                comboBox1.SelectedIndex = 1;
            }
            else
            {
                comboBox1.SelectedIndex = 0;
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            
        }

        private void update_dp_Load(object sender, EventArgs e)
        {

        }

        private void comboBox6_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button8_Click_1(object sender, EventArgs e)
        {
            
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            this.Close();
        }

        private void tableLayoutPanel12_Leave(object sender, EventArgs e)
        {
            
        }

        private void update_dp_Leave(object sender, EventArgs e)
        {

        }

        private void update_dp_Deactivate(object sender, EventArgs e)
        {
           // this.Close();
        }

        private void button8_Click_2(object sender, EventArgs e)
        {
            string loot = textBox6.Text;
            string sbdsb = comboBox6.Text;
            string did = label18.Text;
            string sql_update = "UPDATE tblSBDispatch SET LotNo='" + loot + "', SBDSB='" + sbdsb + "',EID='" + textBox1.Text + "',isByHand="+comboBox1.SelectedIndex+" WHERE DID=" + eid;
            OleDbConnection conn;
            var DBPath = Application.StartupPath + "\\db.mdb";
            conn = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=db.mdb;Jet OLEDB:Database Password=qlty;");
            conn.Open();
            OleDbCommand cmdUpdate = new OleDbCommand();
            cmdUpdate.Parameters.Clear();
            cmdUpdate.CommandText = sql_update;
            cmdUpdate.CommandType = CommandType.Text;
            cmdUpdate.Parameters.Add("@LIKeys", OleDbType.LongVarChar);
            cmdUpdate.Connection = conn;
            cmdUpdate.ExecuteNonQuery();
            conn.Close();
            label2.Text = "Information Sucessfully Updated";
            msg msg = new msg();
            msg.Show();
            this.Hide();
        }

        private void button8_MouseEnter(object sender, EventArgs e)
        {
            button8.ForeColor = Color.Transparent;
            button8.BackColor = Color.FromArgb(0, 50, 51);
            button8.FlatAppearance.BorderColor = Color.FromArgb(0, 50, 51);
        }

        private void button8_MouseLeave(object sender, EventArgs e)
        {
            button8.ForeColor = Color.FromArgb(0, 50, 51);
            button8.BackColor = Color.Transparent;            
        }
    }
}
