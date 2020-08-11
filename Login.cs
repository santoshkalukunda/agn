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
    public partial class Login : Form
    {
        public Login()
        {
            InitializeComponent();
        }

        private void label4_Click(object sender, EventArgs e)
        {
            //label4.Hide();

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            if (textBox1.Text == "")
            {
                label4.Show();
            }
            else
            {
                label4.Hide();
            }
            check();
        }

        private void Login_Load(object sender, EventArgs e)
        {

        }
        public  void check()
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
                        textBox1.Clear();
                        //MessageBox.Show("Login Success as Username: " + textBox1.Text);
                        Farwestern_Regional_AG_Nepal fr = new Farwestern_Regional_AG_Nepal();
                        fr.Show();
                        this.Hide();

                    }

                    else
                    {
                        //MessageBox.Show("Incorrect password !");
                        //textBox1.Clear();

                    }
                }
            }


            catch 
            {

                ///MessageBox.Show(ex.ToString());
            }
        }
        private void button1_Click(object sender, EventArgs e)
        {



          

        }

        private void button1_MouseClick(object sender, MouseEventArgs e)
        {
            
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (textBox1.Text == "")
            {
                label4.Show();
            }
            else
            {
                label4.Hide();
            }
            check();
        }
    }
}
