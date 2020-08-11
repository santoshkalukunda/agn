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
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
            dateTimePicker1.Value = DateTime.Now;
            dateTimePicker1.Format = DateTimePickerFormat.Custom;
            dateTimePicker1.CustomFormat = "yyyy/MM";
           // OleDbConnection conn = new OleDbConnection();
            //string strP = @"Provider = Microsoft.ACE.OLEDB.12.0; Data Source =db.accdb";
            //conn.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=db.accdb";

        }
      

        private void Form2_Load(object sender, EventArgs e)
        {
            greedview();
        }

        private void button1_Click(object sender, EventArgs e)
        {
           OleDbConnection conn = new OleDbConnection();
            conn.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=db.accdb";
            conn.Open();
            string sqlstr= "CREATE TABLE [" +dateTimePicker1.Text + "] ([ID] Counter Primary Key,[SN] Text(30), [Date] Text(30), [Particulars] Text(30), [VCh_Type] Text(30), [JF] Text(30), [Debit] Number, [Credit] Number, [Comments] Text(30))";
            
            OleDbCommand cmmd = new OleDbCommand(sqlstr, conn);
           // cmmd.CommandText ;
            if (conn.State == ConnectionState.Open)
            {
                try
                {
                    cmmd.ExecuteNonQuery();
                    MessageBox.Show("Adding Success New Journal Entries month of '"+dateTimePicker1.Text+"' !", "Information",MessageBoxButtons.OK, MessageBoxIcon.Information);
                    conn.Close();

                    OleDbConnection con = new OleDbConnection();
                    con.ConnectionString = @"Provider = Microsoft.ACE.OLEDB.12.0; Data Source =db.accdb";
                    OleDbCommand cmd = con.CreateCommand();
                    con.Open();
                    cmd.CommandText = "Insert into Vaucharlist (Month_of) Values('" + dateTimePicker1.Text + "')";
                    cmd.Connection = con;
                    try
                    {
                        cmd.ExecuteNonQuery();
                        con.Close();
                        greedview();
                        //comboBox1.Text = "";
                    }
                    catch (Exception)
                    {

                    }
                }
                catch (OleDbException expe)
                {
                    MessageBox.Show(expe.Message);
                    conn.Close();
                }
            }
            else
            {
                MessageBox.Show("Error!");
            }

           

        }
        public void greedview()
        {

            string strProvider = @"Provider = Microsoft.ACE.OLEDB.12.0; Data Source = db.accdb";
            string strSql = "Select Month_of from Vaucharlist order by ID desc";
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

            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            OleDbConnection conn = new OleDbConnection();
            conn.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=db.accdb";
            conn.Open();
            string sqlstr = "Drop table [" + dateTimePicker1.Text + "]";
            

            OleDbCommand cmmd = new OleDbCommand(sqlstr, conn);
           
            // cmmd.CommandText ;
            if (conn.State == ConnectionState.Open)
            {
                try
                {
                    
                   DialogResult result= MessageBox.Show(" Do you want to delete the table '"+dateTimePicker1.Text+"' !", "Confirmation",MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                    if(result==DialogResult.Yes)
                    {
                        
                        cmmd.ExecuteNonQuery();
                       
                        
                    }
                    conn.Close();
                }
                catch (OleDbException expe)
                {
                    MessageBox.Show(expe.Message);
                    conn.Close();
                }
            }
            else
            {
                MessageBox.Show("Error!");
            }

            OleDbConnection con = new OleDbConnection();
            con.ConnectionString = @"Provider = Microsoft.ACE.OLEDB.12.0; Data Source =db.accdb";
            OleDbCommand cmd = con.CreateCommand();
            con.Open();

            cmd.CommandText = "Delete from Vaucharlist where Month_of='" +dateTimePicker1.Text + "'";
            cmd.Connection = con;
            try
            {
               
                {
                    cmd.ExecuteNonQuery();
                    con.Close();
                    greedview();
                    
                   
                   
                }
            }
            catch (Exception)
            {

            }

        }
        public static string tb;
        private void button3_Click(object sender, EventArgs e)
        {
         
            Form2.tb = dateTimePicker1.Text;
            
            Form3 frm3 = new Form3();
           
            frm3.Show();
            this.Hide();
            

        }

        private void button4_Click(object sender, EventArgs e)
        {
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                dateTimePicker1.Text = dataGridView1.SelectedRows[0].Cells[0].Value.ToString();
            }
            catch
            {

            }
        }
    }
}
