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
    public partial class frmDelete : Form
    {
        public frmDelete()
        {
            InitializeComponent();
        }

        private void frmDelete_Load(object sender, EventArgs e)
        {
            label4.Text = Form2.tb;
            greedview();
            label4.Hide();
        }

        public void greedview()
        {

            string strProvider = @"Provider = Microsoft.ACE.OLEDB.12.0; Data Source = db.accdb";
            string strSql = "Select [SN],[Date], [Particulars] from [" + label4.Text + "] order by ID asc ";
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

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
          
            if (textBox1.Text != "")
            {
              
                    OleDbConnection con = new OleDbConnection();
                    con.ConnectionString = @"Provider = Microsoft.ACE.OLEDB.12.0; Data Source =db.accdb";

                    OleDbDataAdapter da = new OleDbDataAdapter("select [SN],[Date], [Particulars] from [" + label4.Text + "] where SN like '" + textBox1.Text + "%'", con);
                    DataTable dt = new DataTable();
                    da.Fill(dt);
                    dataGridView1.DataSource = dt;
                    

                }
            
            else
            {
                textBox2.Clear();
                textBox3.Clear();
                textBox4.Clear();
            }
           
            }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
           
        }

        private void dataGridView1_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try {
                textBox1.Text = dataGridView1.SelectedRows[0].Cells[0].Value.ToString();
                textBox2.Text = dataGridView1.Rows[0].Cells[2].Value.ToString();
                textBox3.Text = dataGridView1.Rows[1].Cells[2].Value.ToString();
                if (textBox1.Text == dataGridView1.Rows[2].Cells[0].Value.ToString())
                {
                    textBox4.Text = dataGridView1.Rows[2].Cells[2].Value.ToString();
                }
            }
            catch
            {

            }
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
          
                    textBox2.Clear();
                    textBox3.Clear();
                    textBox4.Clear();
             
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OleDbConnection con = new OleDbConnection();
            con.ConnectionString = @"Provider = Microsoft.ACE.OLEDB.12.0; Data Source =db.accdb";
            OleDbCommand cmd = con.CreateCommand();
            con.Open();

            cmd.CommandText = "Delete from [" + label4.Text + "] where [SN]='" + textBox1.Text + "'";
            cmd.Connection = con;
            try
            {
                DialogResult rs = MessageBox.Show("Are You sure Delete", "Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (rs == DialogResult.Yes)
                {
                    cmd.ExecuteNonQuery();
                    
                }
            }
            catch
            {

            }
            try
            {
                cmd.CommandText = "Delete from [" + textBox2.Text + "] where [SN]='" + textBox1.Text + "' OR [SN_]='" + textBox1.Text + "'";
                cmd.Connection = con;
                cmd.ExecuteNonQuery();
               
            }
            catch
            {
                
               
            }
            
                try
                {
                    cmd.CommandText = "Delete from [" + textBox4.Text + "] where [SN]='" + textBox1.Text + "'or [SN_]='" + textBox1.Text + "'";
                    cmd.Connection = con;
                    cmd.ExecuteNonQuery();
                    
                }
                catch
                {

                }
            
            try
            {
                cmd.CommandText = "Delete from [" + textBox3.Text + "] where [SN]='" + textBox1.Text + "'or [SN_]='" + textBox1.Text + "'";
                cmd.Connection = con;
                cmd.ExecuteNonQuery();
                
            }
            catch
            {
                
                
            }
            finally
            {
                con.Close();
            }

            textBox1.Clear();
            textBox2.Clear();
            textBox3.Clear();
            textBox4.Clear();
            greedview();
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Form3 fr = new Form3();
            fr.Show();
            this.Hide();

        }
    }
}
