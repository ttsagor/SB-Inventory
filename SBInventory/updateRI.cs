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
    public partial class updateRI : Form
    {
        string DSB = "";
        string loot = "";
        string eid = "";
        Form f1;
        public updateRI(DataTable dt, string par_selected_index, string par_loot, string par_eid, Form form1)
        {
            InitializeComponent();
            Form1.update_reactive = true;
            comboBox1.SelectedIndex = 0;
            foreach (DataRow row in dt.Rows)
            {
                ComboboxItem item = new ComboboxItem();
                item.Text = row[0].ToString();
                item.Value = row[0].ToString();
            }
            DSB = par_selected_index;
            loot = par_loot;
            eid = par_eid;
            if (loot.Trim() == "NEGATIVE")
            { 
                comboBox1.SelectedIndex = 1;
            }
            else if (loot.Trim() == "POSITIVE")
            {
                comboBox1.SelectedIndex = 0;
            }
            //comboBox1.Text = loot;
            f1 = form1;
        }

        private void button8_Click(object sender, EventArgs e)
        {
           
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }        
        private void button8_Click_1(object sender, EventArgs e)
        {
            string sts = "";
            if (comboBox1.Text == "POSITIVE")
            {
                sts = "P";
            }
            else if (comboBox1.Text == "NEGATIVE")
            {
                sts = "N";
            }

            string sql_update = "UPDATE tblSBReceive SET STS='" + sts + "' WHERE RID=" + eid;
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
           // label18.Text = "Information Sucessfully Updated";
            msg msg = new msg();
            msg.Show();
            this.Hide();
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
          this.Close();
        }

        private void tableLayoutPanel12_MouseClick(object sender, MouseEventArgs e)
        {

        }

        private void updateRI_Deactivate(object sender, EventArgs e)
        {
            //this.Close();
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
