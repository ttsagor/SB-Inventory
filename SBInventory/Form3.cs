using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Drawing.Drawing2D;
namespace SBInventory
{
    public partial class Form3 : Form
    {
        public Form3()
        {
            InitializeComponent();
            DoubleBuffered = true;
        }
        OleDbConnection conn;
        private void button1_Click(object sender, EventArgs e)
        {
            var DBPath = Application.StartupPath + "\\db.mdb";
            conn = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=db.mdb;Jet OLEDB:Database Password=qlty;");
            conn.Open();
            string sb_name = textBox1.Text;
            Form1.sb_name_static = sb_name;
            using (DataTable dt = new DataTable())
            {               
                String sql = "DELETE * FROM sb_name_table;";
                using (OleDbDataAdapter adapter = new OleDbDataAdapter(sql, conn))
                {
                    adapter.Fill(dt);
                }

                sql = "INSERT INTO sb_name_table (ID,sb_name) VALUES ('1','" + sb_name + "');";
                using (OleDbDataAdapter adapter = new OleDbDataAdapter(sql, conn))
                {
                    adapter.Fill(dt);
                }
                conn.Close();
               Form f2 = new Form2();
               f2.Show();
               this.Hide();
            }
            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        private void button2_MouseEnter(object sender, EventArgs e)
        {
            button2.ForeColor = Color.Transparent;
            button2.BackColor = Color.Maroon;           
            button2.FlatAppearance.BorderColor = Color.Maroon;
        }

        private void button3_MouseLeave(object sender, EventArgs e)
        {
            button2.BackColor = Color.Transparent;
            button2.ForeColor = Color.Maroon;
        }

        private void button2_MouseLeave(object sender, EventArgs e)
        {
            button2.BackColor = Color.Transparent;
            button2.ForeColor = Color.Maroon;
           // button2.FlatAppearance.BorderColor = Color.Transparent;
        }

        private void button3_MouseEnter(object sender, EventArgs e)
        {
            //button3.BackColor = system;
        }

        private void button1_MouseEnter(object sender, EventArgs e)
        {
            button1.ForeColor = Color.Transparent;
            button1.BackColor = Color.FromArgb(0,51,50);
            button1.FlatAppearance.BorderColor = Color.FromArgb(0, 51, 50);
        }

        private void button1_MouseLeave(object sender, EventArgs e)
        {
            button1.BackColor = Color.Transparent;
            button1.ForeColor = Color.FromArgb(0, 51, 50);
        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                button1_Click(sender, e);
            }
        }
    }
}
