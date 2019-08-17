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
    
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            DataTable dtDSB;
            OleDbConnection conn;

            var DBPath = Application.StartupPath + "\\db.mdb";
            conn = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=db.mdb;Jet OLEDB:Database Password=qlty;");
            conn.Open();
            Boolean isTrial = false;

            

            using (DataTable dt_limit = new DataTable())
            {
                String sql_limit = "SELECT * FROM tblSBDispatch";
                using (OleDbDataAdapter adapter_limit = new OleDbDataAdapter(sql_limit, conn))
                {
                    adapter_limit.Fill(dt_limit);
                   // dtDSB = dt_limit;
                }

                int row_limit = 0;
                foreach (DataRow row in dt_limit.Rows)
                {
                    Form1.sb_name_static = row[0].ToString();
                    row_limit++;
                }
                //MessageBox.Show(dt_limit.Rows.Count.ToString());
                if (isTrial && dt_limit.Rows.Count >= 20000)
                {
                    MessageBox.Show("Trial Period Expired.Please Purchase Full Version.Contact 01914811118");
                    Application.Exit();
                }
                else
                {
                    using (DataTable dt = new DataTable())
                    {
                        String sql = "SELECT sb_name FROM sb_name_table";
                        using (OleDbDataAdapter adapter = new OleDbDataAdapter(sql, conn))
                        {
                            adapter.Fill(dt);
                            dtDSB = dt;
                        }

                        int sb_counter = 0;
                        foreach (DataRow row in dt.Rows)
                        {
                            Form1.sb_name_static = row[0].ToString();

                            sb_counter++;
                        }
                        if (sb_counter > 0)
                        {
                            using (DataTable dt1 = new DataTable())
                            {
                                int sb_office_counter = 0;

                                string sql1 = "SELECT SBDSB FROM tblSBDSB WHERE active_stat=1";
                                using (OleDbDataAdapter adapter1 = new OleDbDataAdapter(sql1, conn))
                                {
                                    adapter1.Fill(dt1);
                                }

                                foreach (DataRow row in dt1.Rows)
                                {
                                    sb_office_counter++;
                                }
                                if (sb_office_counter > 0)
                                {
                                    Application.EnableVisualStyles();
                                    Application.SetCompatibleTextRenderingDefault(false);
                                    Application.Run(new usercontrol());
                                    //Application.Run(new Form3());
                                }
                                else
                                {
                                    Application.EnableVisualStyles();
                                    Application.SetCompatibleTextRenderingDefault(false);
                                    Application.Run(new Form2());
                                    //Application.Run(new Form3());
                                }
                                //  comboBox5.SelectedIndex = 0;
                            }
                        }
                        else
                        {
                            Application.EnableVisualStyles();
                            Application.SetCompatibleTextRenderingDefault(false);
                            Application.Run(new Form3());
                        }

                    }
                }
            }
            

            
            
            conn.Close();             
        }
    }
}
