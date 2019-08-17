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
    public partial class Form2 : Form
    {
        public Form2()
        {
            DoubleBuffered = true;
            InitializeComponent();            
        }
        DataTable dtDSB;
        OleDbConnection conn;
        private void Form2_Load(object sender, EventArgs e)
        {
            var DBPath = Application.StartupPath + "\\db.mdb";
            conn = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=db.mdb;Jet OLEDB:Database Password=qlty;");
            conn.Open();
            using (DataTable dt = new DataTable())
            {
                // txtSelect.Text:
                // SELECT id, text_col, int_col FROM Table_1
                // or
                // SELECT * FROM Table_1
                //
                // selects all content from table and adds it to datatable binded to datagridview
                String sql = "SELECT SBDSB FROM tblSBDSB";
                using (OleDbDataAdapter adapter = new OleDbDataAdapter(sql, conn))
                {
                    adapter.Fill(dt);
                    dtDSB = dt;
                }


                foreach (DataRow row in dt.Rows)
                {
                    listBox1.Items.Add(row[0].ToString());
                }
                
            }
            
            conn.Close();  
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            
            if (listBox1.SelectedItem!=null && !listBox2.Items.Contains(listBox1.SelectedItem))
            {
                listBox2.Items.Add(listBox1.SelectedItem);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (listBox2.SelectedIndex != -1)
            {

                listBox2.Items.Remove(listBox2.SelectedItem);
            }
            
        }

        private void button3_Click(object sender, EventArgs e)
        {

            var DBPath = Application.StartupPath + "\\db.mdb";
            conn = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=db.mdb;Jet OLEDB:Database Password=qlty;");
            conn.Open();
            using (DataTable dt = new DataTable())
            {
                String sql = "UPDATE tblSBDSB SET active_stat='0'";
                using (OleDbDataAdapter adapter = new OleDbDataAdapter(sql, conn))
                {
                    adapter.Fill(dt);
                }

                foreach (string i in listBox2.Items)
                {

                    sql = "UPDATE tblSBDSB SET active_stat=1 WHERE SBDSB='" + i + "'";
                    //MessageBox.Show(sql);
                    using (OleDbDataAdapter adapter = new OleDbDataAdapter(sql, conn))
                    {
                        adapter.Fill(dt);
                    }
                }
            }            
            conn.Close();
            Form f1 = new Form1();
            this.Hide();
            f1.Show();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        private void button4_MouseEnter(object sender, EventArgs e)
        {
            button4.ForeColor = Color.Transparent;
            button4.BackColor = Color.Maroon;
            button4.FlatAppearance.BorderColor = Color.Maroon;
        }

        private void button4_MouseLeave(object sender, EventArgs e)
        {
            button4.BackColor = Color.Transparent;
            button4.ForeColor = Color.Maroon;
        }

        private void button3_MouseEnter(object sender, EventArgs e)
        {
            button3.ForeColor = Color.Transparent;
            button3.BackColor = Color.FromArgb(0, 51, 50);
            button3.FlatAppearance.BorderColor = Color.FromArgb(0, 51, 50);
        }

        private void button3_MouseLeave(object sender, EventArgs e)
        {
            button3.BackColor = Color.Transparent;
            button3.ForeColor = Color.FromArgb(0, 51, 50);
        }

        private void button2_MouseEnter(object sender, EventArgs e)
        {
            button2.ForeColor = Color.Transparent;
            button2.BackColor = Color.Maroon;
            button2.FlatAppearance.BorderColor = Color.Maroon;
        }

        private void button2_MouseLeave(object sender, EventArgs e)
        {
            button2.BackColor = Color.Transparent;
            button2.ForeColor = Color.Maroon;
        }

        private void button1_MouseEnter(object sender, EventArgs e)
        {
            button1.ForeColor = Color.Transparent;
            button1.BackColor = Color.DarkGreen;
            button1.FlatAppearance.BorderColor = Color.DarkGreen;
        }

        private void button1_MouseLeave(object sender, EventArgs e)
        {
            button1.ForeColor = Color.DarkGreen;
            button1.BackColor = Color.Transparent;
        }
    }
}
