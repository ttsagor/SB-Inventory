using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;

namespace SBInventory
{
    public partial class usercontrol : Form
    {
        OleDbConnection conn;
        public static String username = "";
        public static String password = "";
        public static String auto_submit_eid = "0";

        public usercontrol()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            submit();
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            
        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                submit();
            }
        }

        private void textBox2_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                submit();
            }
        }

        void submit()
        {
            username = textBox1.Text;
            password = textBox2.Text;

            if (username.Equals("") || password.Equals(""))
            {
                MessageBox.Show("Please Input Username and Password");
                username = "";
                password = "";
                return;
            }

            DataTable dtDSB;
            OleDbConnection conn;

            var DBPath = Application.StartupPath + "\\db.mdb";
            conn = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=db.mdb;Jet OLEDB:Database Password=qlty;");
            conn.Open();

            using (DataTable dt_limit = new DataTable())
            {
                String sql_limit = "SELECT * FROM usercontrol WHERE username='" + username + "' AND password='" + password + "'";
                using (OleDbDataAdapter adapter_limit = new OleDbDataAdapter(sql_limit, conn))
                {
                    adapter_limit.Fill(dt_limit);
                }

                if (dt_limit.Rows.Count > 0)
                {
                    foreach (DataRow row in dt_limit.Rows)
                    {
                        auto_submit_eid = row[3].ToString();
                    }

                   // MessageBox.Show(auto_submit_eid);

                    Form f1 = new Form1();
                    this.Hide();
                    f1.Show();
                }
                else
                {
                    MessageBox.Show("Username or Password not Matched");
                }
            }
            conn.Close();
        }
    }
}
