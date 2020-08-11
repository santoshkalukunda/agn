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
    public partial class frmVaucherMd : Form
    {
        public frmVaucherMd()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
           
        }
        public void greedview()
        {

            string strProvider = @"Provider = Microsoft.ACE.OLEDB.12.0; Data Source = db.accdb";
            string strSql = "Select * from [" + label3.Text + "]";
            OleDbConnection con = new OleDbConnection(strProvider);
            OleDbCommand cmd = new OleDbCommand(strSql, con);
            con.Open();
            cmd.CommandType = CommandType.Text;
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            DataTable scores = new DataTable();
            try
            {
                da.Fill(scores);
                dataGridView1.DataSource = scores;
            }
            catch (Exception)
            {
                MessageBox.Show("Month Of '" + label3.Text + "' is Not found ! ", "Not Found", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.Close();
                Form2 frm2 = new Form2();
                frm2.Show();
            }
        }
        public void count()
        {
            label5.Text = (dataGridView1.RowCount).ToString();
        }
      
       
        private void frmParticularMd_Load(object sender, EventArgs e)
        {
            comboBox1.SelectedIndex = 0;
            label3.Text = Form3.tbn;
      
            greedview();
            label3.Hide();
            count();
           

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            if (textBox1.Text != "")
            {
                if (comboBox1.Text == "SN")
                {
                    OleDbConnection con = new OleDbConnection();
                    con.ConnectionString = @"Provider = Microsoft.ACE.OLEDB.12.0; Data Source =db.accdb";

                    OleDbDataAdapter da = new OleDbDataAdapter("select * from [" + label3.Text + "] where SN like '" + textBox1.Text + "%'", con);
                    DataTable dt = new DataTable();
                    da.Fill(dt);
                    dataGridView1.DataSource = dt;
                    count();

                }
                else if (comboBox1.Text == "Date")
                {
                    OleDbConnection con = new OleDbConnection();
                    con.ConnectionString = @"Provider = Microsoft.ACE.OLEDB.12.0; Data Source =db.accdb";

                    OleDbDataAdapter da = new OleDbDataAdapter("select * from [" + label3.Text + "] where Date like '" + textBox1.Text + "%'", con);
                    DataTable dt = new DataTable();
                    da.Fill(dt);
                    dataGridView1.DataSource = dt;
                    count();

                }
                else if (comboBox1.Text == "Particulars")
                {
                    OleDbConnection con = new OleDbConnection();
                    con.ConnectionString = @"Provider = Microsoft.ACE.OLEDB.12.0; Data Source =db.accdb";

                    OleDbDataAdapter da = new OleDbDataAdapter("select * from [" + label3.Text + "] where Particulars like '" + textBox1.Text + "%'", con);
                    DataTable dt = new DataTable();
                    da.Fill(dt);
                    dataGridView1.DataSource = dt;
                    count();

                }
                else if (comboBox1.Text == "VCh_Type")
                {
                    OleDbConnection con = new OleDbConnection();
                    con.ConnectionString = @"Provider = Microsoft.ACE.OLEDB.12.0; Data Source =db.accdb";

                    OleDbDataAdapter da = new OleDbDataAdapter("select * from [" + label3.Text + "] where [VCh_Type] like '" + textBox1.Text + "%'", con);
                    DataTable dt = new DataTable();
                    da.Fill(dt);
                    dataGridView1.DataSource = dt;
                    count();

                }
                else
                {

                    MessageBox.Show("Please select valid Search type ");
                    count();
                }
            }
            else
            {
                greedview();
            }
            count();
        }

        private void button2_Click(object sender, EventArgs e)
        {
           
            
            }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
           
        }

        private void groupBox2_Enter(object sender, EventArgs e)
        {

        }

        private void Edit_Enter(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged_1(object sender, EventArgs e)
        {
           
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
           
        }

       

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            
        }

        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
        }

       
        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Form3 fr3 = new Form3();
            fr3.Show();
            this.Hide();
        
        }
    }
}
