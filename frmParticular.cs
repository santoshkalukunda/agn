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
    public partial class frmParticulars : Form
    {
        public frmParticulars()
        {
            InitializeComponent();
        }

        public void greedview()
        {

            string strProvider = @"Provider = Microsoft.ACE.OLEDB.12.0; Data Source = db.accdb";
            string strSql = "Select Particulars from Pt order by ID desc";
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
        public void Combo()
        {
            string q = "select Particulars from Pt";
            OleDbDataAdapter dt = new OleDbDataAdapter(q, @"Provider = Microsoft.ACE.OLEDB.12.0; Data Source = db.accdb");
            DataTable src = new DataTable();



            try
            {
                dt.Fill(src);

                if (src.Rows.Count >= 1)


                {


                    comboBox1.ValueMember = "Particulars";
                    comboBox1.DisplayMember = "Particulars";
                    comboBox1.DataSource = src;
                    comboBox1.Text = "";

                }
            }
            catch (Exception)
            {
                // MessageBox.Show(ex.ToString());
            }
        }
        private void button1_Click(object sender, EventArgs e)
        {

            if (comboBox1.Text != "")
            {
                
                OleDbConnection con = new OleDbConnection();
                con.ConnectionString = @"Provider = Microsoft.ACE.OLEDB.12.0; Data Source =db.accdb";
                OleDbCommand cmd = con.CreateCommand();
                con.Open();
                cmd.CommandText = "Insert into Pt (Particulars) Values('" + comboBox1.Text + "')";
                cmd.Connection = con;
                try {
                    cmd.ExecuteNonQuery();
                    con.Close();
                    greedview();
                    comboBox1.Text = "";
                }
                catch(Exception)
                {
                    MessageBox.Show("Particular '" + comboBox1.Text + "' alredy in the list");
                }
            }
            else
            {
                MessageBox.Show("Please Enter Particulars");
            }

            //string q = "select Particulars from Pt";
            //OleDbDataAdapter dt = new OleDbDataAdapter(q, @"Provider = Microsoft.ACE.OLEDB.12.0; Data Source = db.accdb");
            //DataTable src = new DataTable();



            //try
            //{
            //    dt.Fill(src);

            //    if (src.Rows.Count >= 1)


            //    {


            //        comboBox1.ValueMember = "Particulars";
            //        comboBox1.DisplayMember = "Particulars";
            //        comboBox1.DataSource = src;
            //        comboBox1.Text = "";

            //    }
            //}
            //catch (Exception ex)
            //{
            //    // MessageBox.Show(ex.ToString());
            //}

        }

        private void frmParticulars_Load(object sender, EventArgs e)
        {
            greedview();
            button2.Hide();
            groupBox1.Hide();
            //button5.Hide();


            //string q = "select Particulars from Pt";
            //OleDbDataAdapter dt = new OleDbDataAdapter(q, @"Provider = Microsoft.ACE.OLEDB.12.0; Data Source = db.accdb");
            //DataTable src = new DataTable();



            //try
            //{
            //    dt.Fill(src);

            //    if (src.Rows.Count >= 1)


            //    {


            //        comboBox1.ValueMember = "Particulars";
            //        comboBox1.DisplayMember = "Particulars";
            //        comboBox1.DataSource = src;
            //        comboBox1.Text = "";

            //    }
            //}
            //catch (Exception ex)
            //{
            //    // MessageBox.Show(ex.ToString());
            //}
        }

        private void comboBox1_Click(object sender, EventArgs e)
        {
            //comboBox1.DroppedDown = true;
        }

        private void comboBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            //if(char.IsLower(Convert.ToChar(comboBox1.Text.Substring(0,1))))
            //{
            //    comboBox1.Text = comboBox1.Text.Replace(comboBox1.Text.Substring(0, 1), comboBox1.Text.ToUpper());
            //    comboBox1.SelectionStart = 2;
            //}
            //comboBox1.DroppedDown = true;
            string strProvider = @"Provider = Microsoft.ACE.OLEDB.12.0; Data Source = db.accdb";
            //string strSql = ;
            OleDbConnection con = new OleDbConnection(strProvider);
            // OleDbCommand cmd = new OleDbCommand(strSql, con);
            con.Open();
            //cmd.CommandType = CommandType.Text;
            OleDbDataAdapter da = new OleDbDataAdapter("Select Particulars from Pt where Particulars like '" + comboBox1.Text + "%'", con);
            DataTable scores = new DataTable();
           
                da.Fill(scores);
                dataGridView1.DataSource = scores;
         
          
        }

        private void button2_Click(object sender, EventArgs e)
        {

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                textBox1.Text = dataGridView1.SelectedRows[0].Cells[1].Value.ToString();
            }
            catch
            {

            }
            
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            greedview();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            OleDbConnection con = new OleDbConnection();
            con.ConnectionString = @"Provider = Microsoft.ACE.OLEDB.12.0; Data Source =db.accdb";
            OleDbCommand cmd = con.CreateCommand();
            con.Open();
            
            cmd.CommandText = "Delete from Pt where Particulars='" + comboBox1.Text + "'";
            cmd.Connection = con;
            try {
                DialogResult rs = MessageBox.Show("Are You sure Delete the Particulars is '" + comboBox1.Text + "'", "Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (rs== DialogResult.Yes)
                {
                    cmd.ExecuteNonQuery();
                    con.Close();
                    greedview();
                    comboBox1.Enabled = true;
                    button1.Show();
                    button3.Show();
                    button2.Hide();
                    groupBox1.Hide();
                    //button5.Hide();
                    comboBox1.Text = "";
                    //Combo();
                }
                
                comboBox1.Text = "";
            }
            catch(Exception)
            {

            }
            

        }

        private void dataGridView1_MouseClick(object sender, MouseEventArgs e)
        {
            try
            {
                comboBox1.Text = dataGridView1.SelectedRows[0].Cells[0].Value.ToString();
                comboBox1.Enabled = false;
                button1.Hide();
                button2.Show();
                groupBox1.Show();
               // button5.Show();
                button3.Hide();

            }
            catch (Exception)
            {
            }


        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (comboBox1.Text != "")
            {
                OleDbConnection con = new OleDbConnection();
                con.ConnectionString = @"Provider = Microsoft.ACE.OLEDB.12.0; Data Source =db.accdb";

                OleDbDataAdapter da = new OleDbDataAdapter("select * from Pt", con);
                DataTable dt = new DataTable();
                da.Fill(dt);
                try
                {

                    if (dt.Rows.Count >= 1)
                    {
                        int i = 0;
                        while (comboBox1.Text != dt.Rows[i].ItemArray[1].ToString())
                        {

                            i++;

                        }
                        comboBox1.Enabled = false;
                        button1.Hide();
                        button2.Show();
                        groupBox1.Show();
                        //button5.Show();
                        button3.Hide();
                    }
                }

                catch (Exception)
                {
                    MessageBox.Show("This Particular is not In the list !");
                }
            }
            else
            {
                MessageBox.Show("Please enter which you want to edit");
            }
            
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (textBox1.Text != "")
            {
                string strProvider = @"Provider = Microsoft.ACE.OLEDB.12.0; Data Source = db.accdb";
                string strSql = "Update Pt set Particulars='" + textBox1.Text + "' where Particulars='" + comboBox1.Text + "'";
                OleDbConnection con = new OleDbConnection(strProvider);
                OleDbCommand cmd = new OleDbCommand(strSql, con);
                con.Open();
                
                    try
                    {
                        cmd.ExecuteNonQuery();
                        MessageBox.Show("Update Successfull !");
                        con.Close();
                        greedview();
                        textBox1.Clear();
                        comboBox1.Text = "";
                        comboBox1.Enabled = true;
                        button2.Hide();
                        groupBox1.Hide();
                        //button5.Hide();
                        button1.Show();
                        button3.Show();
                    }
                    catch (Exception)
                    {
                        MessageBox.Show("The particular is already in the list");
                    }
                
            }
            else
            {
                MessageBox.Show("Please Enter the valid paticulars");
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
          
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Form3 fr2 = new Form3();
            fr2.Show();
            this.Hide();
        }
    }
}
