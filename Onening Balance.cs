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
    public partial class Opening_Balance : Form
    {
        public Opening_Balance()
        {
            InitializeComponent();
        }

        private void Onening_Balance_Load(object sender, EventArgs e)
        {
            comboBox1.SelectedIndex = 0;
            comboBox2.SelectedIndex = 0;
            greedview();
        }

        public void greedview()
        {

            string strProvider = @"Provider = Microsoft.ACE.OLEDB.12.0; Data Source = db.accdb";
            string strSql = "Select * from [" +comboBox1.Text + "] where [ID]=9999";
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
                con.Close();
            }
            catch (Exception)
            {
               
            }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            if((comboBox1.Text=="Cash" || comboBox1.Text=="Bank")&& (comboBox2.Text=="Debit" || comboBox2.Text=="Credit")&& textBox1.Text!="")
            {
                OleDbConnection con = new OleDbConnection();
                con.ConnectionString = @"Provider = Microsoft.ACE.OLEDB.12.0; Data Source =db.accdb";
                OleDbCommand cmd = con.CreateCommand();
                con.Open();

                if (comboBox2.Text=="Debit")
                {
                   
                        
                      
                        try
                        {
                            cmd.CommandText = "Insert into [" +comboBox1.Text + "] ([ID], [SN], [Date], [Particulars], [Amount]) Values('9999','1','" + dateTimePicker1.Text + "','Opening Balance','" + textBox1.Text + "')";
                            cmd.Connection = con;
                            cmd.ExecuteNonQuery();

                        cmd.CommandText = "insert into [Ptlist] ([Particulars]) values ('" + comboBox1.Text + "')";
                        cmd.Connection = con;
                        cmd.ExecuteNonQuery();

                        if (comboBox1.Text == "Cash")
                        {
                            cmd.CommandText = "CREATE TABLE [Opening_Balance_Cash] ([ID] Counter Primary Key, [SN] Text(30), [Date] Text(30), [Particulars] Text(30), [Amount] Number, [SN_] text(30), [Date_] Text(30), [Particulars_] Text(30), [Amount_] Number)";
                            cmd.Connection = con;
                            cmd.ExecuteNonQuery();

                            cmd.CommandText = "Insert into [Opening_Balance_Cash] ([ID], [SN_], [Date_], [Particulars_], [Amount_]) Values('9999','1','" + dateTimePicker1.Text + "','Opening Balance','" + textBox1.Text + "')";
                            cmd.Connection = con;
                            cmd.ExecuteNonQuery();

                            cmd.CommandText = "insert into [Ptlist] ([Particulars]) values ('Opening_Balance_Cash')";
                            cmd.Connection = con;
                            cmd.ExecuteNonQuery();
                        }
                        else
                        {
                            cmd.CommandText = "CREATE TABLE [Opening_Balance_Bank] ([ID] Counter Primary Key, [SN] Text(30), [Date] Text(30), [Particulars] Text(30), [Amount] Number, [SN_] text(30), [Date_] Text(30), [Particulars_] Text(30), [Amount_] Number)";
                            cmd.Connection = con;
                            cmd.ExecuteNonQuery();

                            cmd.CommandText = "Insert into [Opening_Balance_Bank] ([ID], [SN_], [Date_], [Particulars_], [Amount_]) Values('9999','1','" + dateTimePicker1.Text + "','Opening Balance','" + textBox1.Text + "')";
                            cmd.Connection = con;
                            cmd.ExecuteNonQuery();

                            cmd.CommandText = "insert into [Ptlist] ([Particulars]) values ('Opening_Balance_Bank')";
                            cmd.Connection = con;
                            cmd.ExecuteNonQuery();
                        }
                        MessageBox.Show("Success openning balance is insered in " + comboBox1.Text + " amount " + textBox1.Text + " Debit ");
                        textBox1.Clear();
                        greedview();
                        con.Close();
                    }
                        catch (Exception)
                        {

                        try {
                            cmd.CommandText = "CREATE TABLE [" + comboBox1.Text + "] ([ID] Counter Primary Key, [SN] Text(30), [Date] Text(30), [Particulars] Text(30), [Amount] Number, [SN_] text(30), [Date_] Text(30), [Particulars_] Text(30), [Amount_] Number)";
                            cmd.Connection = con;
                            cmd.ExecuteNonQuery();
                            cmd.CommandText = "Insert into [" + comboBox1.Text + "] ([ID], [SN], [Date], [Particulars], [Amount]) Values('9999','1','" + dateTimePicker1.Text + "','Opening Balance','" + textBox1.Text + "')";
                            cmd.Connection = con;
                            cmd.ExecuteNonQuery();
                            cmd.CommandText = "insert into [Ptlist] ([Particulars]) values ('" + comboBox1.Text + "')";
                            cmd.Connection = con;
                            cmd.ExecuteNonQuery();

                            if (comboBox1.Text == "Cash")
                            {
                                cmd.CommandText = "CREATE TABLE [Opening_Balance_Cash] ([ID] Counter Primary Key, [SN] Text(30), [Date] Text(30), [Particulars] Text(30), [Amount] Number, [SN_] text(30), [Date_] Text(30), [Particulars_] Text(30), [Amount_] Number)";
                                cmd.Connection = con;
                                cmd.ExecuteNonQuery();

                                cmd.CommandText = "Insert into [Opening_Balance_Cash] ([ID], [SN_], [Date_], [Particulars_], [Amount_]) Values('9999','1','" + dateTimePicker1.Text + "','Opening Balance','" + textBox1.Text + "')";
                                cmd.Connection = con;
                                cmd.ExecuteNonQuery();

                                cmd.CommandText = "insert into [Ptlist] ([Particulars]) values ('Opening_Balance_Cash')";
                                cmd.Connection = con;
                                cmd.ExecuteNonQuery();
                            }
                            else
                            {
                                cmd.CommandText = "CREATE TABLE [Opening_Balance_Bank] ([ID] Counter Primary Key, [SN] Text(30), [Date] Text(30), [Particulars] Text(30), [Amount] Number, [SN_] text(30), [Date_] Text(30), [Particulars_] Text(30), [Amount_] Number)";
                                cmd.Connection = con;
                                cmd.ExecuteNonQuery();

                                cmd.CommandText = "Insert into [Opening_Balance_Bank] ([ID], [SN_], [Date_], [Particulars_], [Amount_]) Values('9999','1','" + dateTimePicker1.Text + "','Opening Balance','" + textBox1.Text + "')";
                                cmd.Connection = con;
                                cmd.ExecuteNonQuery();

                                cmd.CommandText = "insert into [Ptlist] ([Particulars]) values ('Opening_Balance_Bank')";
                                cmd.Connection = con;
                                cmd.ExecuteNonQuery();
                            }

                            MessageBox.Show("Success openning balance is insered in " + comboBox1.Text + " amount " + textBox1.Text + " Debit ");
                            textBox1.Clear();
                            greedview();
                            con.Close();
                        }
                        catch
                        {
                            //MessageBox.Show("Alredy opening balance is inserted " + comboBox1.Text + "  ");
                            greedview();
                        }
                     }
                }
                else if(comboBox2.Text=="Credit")
                {
                    try
                    {
                        cmd.CommandText = "Insert into [" + comboBox1.Text + "] ([ID], [SN_], [Date_], [Particulars_], [Amount_]) Values('9999','1','" + dateTimePicker1.Text + "','Opening Balance','" + textBox1.Text + "')";
                        cmd.Connection = con;
                        cmd.ExecuteNonQuery();
                        cmd.CommandText = "insert into [Ptlist] ([Particulars]) values ('" + comboBox1.Text + "')";
                        cmd.Connection = con;
                        cmd.ExecuteNonQuery();
                        if (comboBox1.Text == "Cash")
                        {
                            cmd.CommandText = "CREATE TABLE [Opening_Balance_Cash] ([ID] Counter Primary Key, [SN] Text(30), [Date] Text(30), [Particulars] Text(30), [Amount] Number, [SN_] text(30), [Date_] Text(30), [Particulars_] Text(30), [Amount_] Number)";
                            cmd.Connection = con;
                            cmd.ExecuteNonQuery();
                            cmd.CommandText = "Insert into [Opening_Balance_Cash] ([ID], [SN], [Date], [Particulars], [Amount]) Values('9999','1','" + dateTimePicker1.Text + "','Opening Balance','" + textBox1.Text + "')";
                            cmd.Connection = con;
                            cmd.ExecuteNonQuery();

                            cmd.CommandText = "insert into [Ptlist] ([Particulars]) values ('Opening_Balance_Cash')";
                            cmd.Connection = con;
                            cmd.ExecuteNonQuery();
                        }
                        else
                        {
                            cmd.CommandText = "CREATE TABLE [Opening_Balance_Bank] ([ID] Counter Primary Key, [SN] Text(30), [Date] Text(30), [Particulars] Text(30), [Amount] Number, [SN_] text(30), [Date_] Text(30), [Particulars_] Text(30), [Amount_] Number)";
                            cmd.Connection = con;
                            cmd.ExecuteNonQuery();
                            cmd.CommandText = "Insert into [Opening_Balance_Bank] ([ID], [SN], [Date], [Particulars], [Amount]) Values('9999','1','" + dateTimePicker1.Text + "','Opening Balance','" + textBox1.Text + "')";
                            cmd.Connection = con;
                            cmd.ExecuteNonQuery();


                            cmd.CommandText = "insert into [Ptlist] ([Particulars]) values ('Opening_Balance_Bank')";
                            cmd.Connection = con;
                            cmd.ExecuteNonQuery();
                        }
                        con.Close();
                        
                        MessageBox.Show("Success openning balance is insered in " + comboBox1.Text + " amount "+textBox1.Text + " Credit ");
                        textBox1.Clear();
                        greedview();
                    }
                    catch (Exception)
                    {
                        try {
                            cmd.CommandText = "CREATE TABLE [" + comboBox1.Text + "] ([ID] Counter Primary Key, [SN] Text(30), [Date] Text(30), [Particulars] Text(30), [Amount] Number, [SN_] text(30), [Date_] Text(30), [Particulars_] Text(30), [Amount_] Number)";
                            cmd.Connection = con;
                            cmd.ExecuteNonQuery();

                            cmd.CommandText = "Insert into [" + comboBox1.Text + "] ([ID], [SN_], [Date_], [Particulars_], [Amount_]) Values('9999','1','" + dateTimePicker1.Text + "','Opening Balance','" + textBox1.Text + "')";
                            cmd.Connection = con;
                            cmd.ExecuteNonQuery();
                            cmd.CommandText = "insert into [Ptlist] ([Particulars]) values ('" + comboBox1.Text + "')";
                            cmd.Connection = con;
                            cmd.ExecuteNonQuery();

                            if (comboBox1.Text == "Cash")
                            {
                                cmd.CommandText = "CREATE TABLE [Opening_Balance_Cash] ([ID] Counter Primary Key, [SN] Text(30), [Date] Text(30), [Particulars] Text(30), [Amount] Number, [SN_] text(30), [Date_] Text(30), [Particulars_] Text(30), [Amount_] Number)";
                                cmd.Connection = con;
                                cmd.ExecuteNonQuery();
                                cmd.CommandText = "Insert into [Opening_Balance_Cash] ([ID], [SN], [Date], [Particulars], [Amount]) Values('9999','1','" + dateTimePicker1.Text + "','Opening Balance','" + textBox1.Text + "')";
                                cmd.Connection = con;
                                cmd.ExecuteNonQuery();

                                cmd.CommandText = "insert into [Ptlist] ([Particulars]) values ('Opening_Balance_Cash')";
                                cmd.Connection = con;
                                cmd.ExecuteNonQuery();
                            }
                            else
                            {
                                cmd.CommandText = "CREATE TABLE [Opening_Balance_Bank] ([ID] Counter Primary Key, [SN] Text(30), [Date] Text(30), [Particulars] Text(30), [Amount] Number, [SN_] text(30), [Date_] Text(30), [Particulars_] Text(30), [Amount_] Number)";
                                cmd.Connection = con;
                                cmd.ExecuteNonQuery();
                                cmd.CommandText = "Insert into [Opening_Balance_Bank] ([ID], [SN], [Date], [Particulars], [Amount]) Values('9999','1','" + dateTimePicker1.Text + "','Opening Balance','" + textBox1.Text + "')";
                                cmd.Connection = con;
                                cmd.ExecuteNonQuery();


                                cmd.CommandText = "insert into [Ptlist] ([Particulars]) values ('Opening_Balance_Bank')";
                                cmd.Connection = con;
                                cmd.ExecuteNonQuery();
                            }

                            MessageBox.Show("Success openning balance is insered in "+comboBox1.Text+" amount "+textBox1.Text+" Credit ");
                            textBox1.Clear();
                            greedview();
                        }
                        catch
                        {

                            //MessageBox.Show("  Alredy opening balance is inserted in "+comboBox1.Text+" ");
                            greedview();
                        }
                    }
                    
                }

            }
            else
            {
                MessageBox.Show("Please Fill proper data");
            }
            textBox1.Clear();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            greedview();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            OleDbConnection con = new OleDbConnection();
            con.ConnectionString = @"Provider = Microsoft.ACE.OLEDB.12.0; Data Source =db.accdb";
            OleDbCommand cmd = con.CreateCommand();
            OleDbCommand cmd1 = con.CreateCommand();
            con.Open();

            cmd.CommandText = "Delete from ["+comboBox1.Text+"] where [ID] =9999";
            if (comboBox1.Text == "Cash")
            {
                cmd1.CommandText = "Drop table [Opening_Balance_Cash]";
                //cmd.CommandText = "Delete from [Ptlist] where [Particulars] =Opening_Balance_Cash";
            }
            else
            {
                cmd1.CommandText = "Drop table [Opening_Balance_Bank]";
                // cmd.CommandText = "Delete from [Ptlist] where [Particulars] =Opening_Balance_Cash";

            }
            cmd.Connection = con;
            cmd1.Connection = con;
            try
            {
                DialogResult rs = MessageBox.Show("Are You sure Delete the Opening Balance in '" + comboBox1.Text + "'", "Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (rs == DialogResult.Yes)
                {
                    cmd.ExecuteNonQuery();
                    cmd1.ExecuteNonQuery();
                    con.Close();
                    greedview();
                  
                }

                
            }
            catch (Exception)
            {
                greedview();
                //MessageBox.Show("Not found");
            }
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {
                e.Handled = true;

            }
            if ((e.KeyChar == '.') && ((sender as TextBox).Text.IndexOf('.') > -1))
            {
                e.Handled = true;
            }
        }
    }
}
