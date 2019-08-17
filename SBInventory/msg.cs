using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace SBInventory
{
    public partial class msg : Form
    {
        public msg()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button1_MouseEnter(object sender, EventArgs e)
        {
           
        }

        private void button1_MouseHover(object sender, EventArgs e)
        {
            
        }

        private void button1_MouseLeave(object sender, EventArgs e)
        {
            
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button1_MouseEnter_1(object sender, EventArgs e)
        {
            button1.ForeColor = Color.Transparent;
            button1.BackColor = Color.FromArgb(0,51,50);
            button1.FlatAppearance.BorderColor = Color.FromArgb(0, 51, 50);
        }

        private void button1_MouseLeave_1(object sender, EventArgs e)
        {
            button1.ForeColor = Color.FromArgb(0, 51, 50);
            button1.BackColor = Color.Transparent;
        }

        private void msg_Deactivate(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
