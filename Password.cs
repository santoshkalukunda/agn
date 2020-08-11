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

namespace DahitGroups
{
    public partial class Password : Form
    {
        public Password()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox2.Text == textBox3.Text)
            {
                OleDbConnection con = new OleDbConnection();
                con.ConnectionString = @"Provider = Microsoft.ACE.OLEDB.12.0; Data Source =db.accdb";
               
                //OleDbDataAdapter da = new OleDbDataAdapter("select * from [Login] where [Password]="+textBox1.Text+"", con);
                //DataTable dt = new DataTable();
                //da.Fill(dt);
                try
                {
                    OleDbDataAdapter da = new OleDbDataAdapter("select * from [Login]", con);
                    DataTable dt = new DataTable();
                    da.Fill(dt);

                    if (dt.Rows.Count >= 1)
                    {
                        if (dt.Rows[0].ItemArray[0].ToString() == textBox1.Text)
                        {
                            OleDbCommand cmd = con.CreateCommand();
                            con.Open();
                            //string strSql = "Update Pt set Particulars='" + textBox1.Text + "' where Particulars='" + comboBox1.Text + "'";

                            cmd.CommandText = "Update [Login] set [Password]='" + textBox2.Text + "' where [ID]='1'";
                            cmd.Connection = con;
                            cmd.ExecuteNonQuery();
                            MessageBox.Show("Success to channge Password");
                            textBox1.Clear();
                            textBox2.Clear();
                            textBox3.Clear();

                        }
                        else
                        {
                            MessageBox.Show("Current Password is incurect !");
                            textBox1.Clear();
                            textBox2.Clear();
                            textBox3.Clear();
                        }
                    }
                }


                catch(Exception ex)
                {

                    MessageBox.Show("Current Password is incurect !");
                    textBox1.Clear();
                    textBox2.Clear();
                    textBox3.Clear();
                }
            }
            else
            {
                MessageBox.Show("The new password and re-enter password is not Matched ");
                textBox1.Clear();
                textBox2.Clear();
                textBox3.Clear();
            }
        }
    }
}
