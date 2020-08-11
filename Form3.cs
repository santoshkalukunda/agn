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
    public partial class Form3 : Form
    {
        public Form3()
        {
            InitializeComponent();
            dateTimePicker1.Value = DateTime.Now;
            dateTimePicker1.Format = DateTimePickerFormat.Custom;
            dateTimePicker1.CustomFormat = "yyyy/MM/dd";
            


        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox1.SelectedIndex == 2)
            {
                comboBox3.Enabled = false;
                textBox2.Enabled = false;

            }
            else
            {
                comboBox3.Enabled = true;
                if (comboBox3.SelectedIndex == 2)
                {
                    textBox2.Enabled = true;
                }
            }

        }

        public void greedview()
        {

            string strProvider = @"Provider = Microsoft.ACE.OLEDB.12.0; Data Source = db.accdb";
            string strSql = "Select * from [" + label5.Text + "] order by ID desc";
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
                MessageBox.Show("Month Of '" + label5.Text + "' is Not found ! ", "Not Found", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.Close();
                Form2 frm2 = new Form2();
                frm2.Show();
            }
        }

        public void dtotal()
        {
            double sum = 0;
            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                if (Convert.ToString(dataGridView1.Rows[i].Cells[6].Value.ToString()) == "")
                {
                    continue;
                }
                sum = sum + Convert.ToDouble(dataGridView1.Rows[i].Cells[6].Value.ToString());
            }
            label14.Text = sum.ToString();
        }
        public void ctotal()
        {
            double sum = 0;
            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                if (Convert.ToString(dataGridView1.Rows[i].Cells[7].Value.ToString()) == "")
                {
                    continue;
                }
                sum = sum + Convert.ToDouble(dataGridView1.Rows[i].Cells[7].Value.ToString());
            }
            label15.Text = sum.ToString();
        }

        private void comboBox1_Click(object sender, EventArgs e)
        {
            comboBox1.DroppedDown = true;
        }
        public static string tbn;
        private void Form3_Load(object sender, EventArgs e)
        {
            label5.Text = Form2.tb;
            comboBox1.SelectedIndex = 0;
            comboBox3.SelectedIndex = 0;
            greedview();
            label18.Hide();
            dtotal();
            ctotal();
            label17.Hide();

             totalrow = dataGridView1.Rows.Count;


            string q = "select Particulars from Pt";
            OleDbDataAdapter dt = new OleDbDataAdapter(q, @"Provider = Microsoft.ACE.OLEDB.12.0; Data Source = db.accdb");

            DataTable src = new DataTable();



            try
            {
                dt.Fill(src);

                if (src.Rows.Count >= 1)


                {


                    comboBox2.ValueMember = "Particulars";
                    comboBox2.DisplayMember = "Particulars";
                    comboBox2.DataSource = src;
                    comboBox2.Text = "";

                }
            }
            catch (Exception)
            {
                // MessageBox.Show(ex.ToString());
            }

            
        }

        public string createaccout = "[ID] Counter Primary Key, [SN] Text(30), [Date] Text(30), [Particulars] Text(30), [Amount] Number, [SN_] text(30), [Date_] Text(30), [Particulars_] Text(30), [Amount_] Number";
      

        public void payment()
        {
            //string debit = "";
            OleDbConnection con = new OleDbConnection();
            con.ConnectionString = @"Provider = Microsoft.ACE.OLEDB.12.0; Data Source =db.accdb";
            OleDbCommand cmd = con.CreateCommand();
            con.Open();

            if (comboBox3.SelectedIndex == 0)// cash
            {
                textBox2.Enabled = false;
                textBox2.Text = textBox1.Text;
                try
                {
                    
                    cmd.CommandText = "Insert into [" + label5.Text + "] ([SN], [Date], [Particulars], [VCh_Type], [Debit], [Comments]) Values('" + label18.Text + "','" + dateTimePicker1.Text + "','" + comboBox2.Text + "','" + comboBox1.Text + "','" + textBox1.Text + "','" + textBox4.Text + "')";
                    cmd.Connection = con;
                    cmd.ExecuteNonQuery();
                    cmd.CommandText = "Insert into [" + label5.Text + "] ([SN], [Particulars], [VCh_Type], [Credit]) Values('" + label18.Text + "','" + comboBox3.Text + "','" + comboBox1.Text + "','" + textBox2.Text + "')";
                    cmd.Connection = con;
                    cmd.ExecuteNonQuery();

                    try
                    {
                        cmd.CommandText = "Insert into [" + comboBox2.Text + "] ([SN], [Date], [Particulars], [Amount]) Values('" + label18.Text + "','" + dateTimePicker1.Text + "','" + comboBox3.Text + "','" + textBox1.Text + "')";
                        cmd.Connection = con;
                        cmd.ExecuteNonQuery();
                    }
                    catch (Exception)
                    {

                        cmd.CommandText = "CREATE TABLE [" + comboBox2.Text + "] ("+createaccout+")";
                        cmd.Connection = con;
                        cmd.ExecuteNonQuery();
                        cmd.CommandText = "Insert into [" + comboBox2.Text + "] ([SN], [Date], [Particulars], [Amount]) Values('" + label18.Text + "','" + dateTimePicker1.Text + "','" + comboBox3.Text + "','" + textBox1.Text + "')";
                        cmd.Connection = con;
                        cmd.ExecuteNonQuery();

                        cmd.CommandText = "insert into [Ptlist] ([Particulars]) values ('" + comboBox2.Text + "')";
                        cmd.Connection = con;
                        cmd.ExecuteNonQuery();

                    }
                    try {

                        cmd.CommandText = "Insert into [" + comboBox3.Text + "] ([SN_],[Date_], [Particulars_], [Amount_]) Values('" + label18.Text + "','" + dateTimePicker1.Text + "','" + comboBox2.Text + "','" + textBox1.Text + "')";
                        cmd.Connection = con;
                        cmd.ExecuteNonQuery();
                    }
                    catch (Exception)
                    {
                        cmd.CommandText = "CREATE TABLE [" + comboBox3.Text + "] ("+createaccout+")";
                        cmd.ExecuteNonQuery();
                        cmd.CommandText = "Insert into [" + comboBox3.Text + "] ([SN_], [Date_], [Particulars_], [Amount_]) Values('" + label18.Text + "','" + dateTimePicker1.Text + "','" + comboBox2.Text + "','" + textBox1.Text + "')";
                        cmd.Connection = con;
                        cmd.ExecuteNonQuery();

                        cmd.CommandText = "insert into [Ptlist] ([Particulars]) values ('" + comboBox3.Text + "')";
                        cmd.Connection = con;
                        cmd.ExecuteNonQuery();

                    }
                    con.Close();
                    greedview();

                    textBox1.Clear();
                    textBox2.Clear();
                    textBox3.Clear();
                    textBox4.Clear();
                    textBox2.Enabled = false;
                    comboBox2.Text = "";


                }
                catch (Exception)
                {
                    //MessageBox.Show(ex.ToString());


                    con.Close();
                    greedview();

                    textBox1.Clear();
                    textBox2.Clear();
                    textBox3.Clear();
                    textBox4.Clear();
                    textBox2.Enabled = false;
                    comboBox2.Text = "";

                }
            }

            else if (comboBox3.Text == "Bank") //bank
            {
                //textBox3.Text = textBox1.Text;
                textBox2.Enabled = false;
                try
                {
                    cmd.CommandText = "Insert into [" + label5.Text + "] ([SN], [Date], [Particulars], [VCh_Type], [Debit], [Comments]) Values('" + label18.Text + "','" + dateTimePicker1.Text + "','" + comboBox2.Text + "','" + comboBox1.Text + "', '" + textBox1.Text + "','" + textBox4.Text + "')";
                    cmd.Connection = con;
                    cmd.ExecuteNonQuery();
                    cmd.CommandText = "Insert into [" + label5.Text + "] ([SN], [Particulars], [VCh_Type], [Credit]) Values('" + label18.Text + "','" + comboBox3.Text + "','" + comboBox1.Text + "','" + textBox1.Text + "')";
                    cmd.Connection = con;
                    cmd.ExecuteNonQuery();


                    try
                    {
                        cmd.CommandText = "Insert into [" + comboBox2.Text + "] ([SN], [Date], [Particulars], [Amount]) Values('" + label18.Text + "','" + dateTimePicker1.Text + "','" + comboBox3.Text + "','" + textBox1.Text + "')";
                        cmd.Connection = con;
                        cmd.ExecuteNonQuery();
                    }
                    catch (Exception)
                    {

                        cmd.CommandText = "CREATE TABLE [" + comboBox2.Text + "] ("+createaccout+")";
                        cmd.Connection = con;
                        cmd.ExecuteNonQuery();
                        cmd.CommandText = "Insert into [" + comboBox2.Text + "] ([SN], [Date], [Particulars], [Amount]) Values('" + label18.Text + "','" + dateTimePicker1.Text + "','" + comboBox3.Text + "','" + textBox1.Text + "')";
                        cmd.Connection = con;
                        cmd.ExecuteNonQuery();

                        cmd.CommandText = "insert into [Ptlist] ([Particulars]) values ('" + comboBox2.Text + "')";
                        cmd.Connection = con;
                        cmd.ExecuteNonQuery();
                    }
                    try
                    {

                        cmd.CommandText = "Insert into [" + comboBox3.Text + "] ([SN_], [Date_], [Particulars_], [Amount_]) Values('" + label18.Text + "','" + dateTimePicker1.Text + "','" + comboBox2.Text + "','" + textBox1.Text + "')";
                        cmd.Connection = con;
                        cmd.ExecuteNonQuery();
                    }
                    catch (Exception)
                    {
                        cmd.CommandText = "CREATE TABLE [" + comboBox3.Text + "] ("+createaccout+")";
                        cmd.Connection = con;
                        cmd.ExecuteNonQuery();
                        cmd.CommandText = "Insert into [" + comboBox3.Text + "] ([SN_], [Date_], [Particulars_], [Amount_]) Values('" + label18.Text + "','" + dateTimePicker1.Text + "','" + comboBox2.Text + "','" + textBox1.Text + "')";
                        cmd.Connection = con;
                        cmd.ExecuteNonQuery();
                        cmd.CommandText = "insert into [Ptlist] ([Particulars]) values ('" + comboBox3.Text + "')";
                        cmd.Connection = con;
                        cmd.ExecuteNonQuery();
                    }
                    con.Close();
                    greedview();

                    textBox1.Clear();
                    textBox2.Clear();
                    textBox3.Clear();
                    textBox4.Clear();
                    textBox2.Enabled = false;
                    comboBox2.Text = "";
                }
                catch (Exception)
                {
                    //MessageBox.Show(ex.ToString());


                    con.Close();
                    greedview();
                    textBox1.Clear();
                    textBox2.Clear();
                    textBox3.Clear();
                    textBox4.Clear();
                    textBox2.Enabled = false;
                    comboBox2.Text = "";
                }
            }


            else if (comboBox3.SelectedIndex == 2)// cash and bank
            {
                if (textBox1.Text != "" && textBox2.Text != "")
                {
                    textBox2.Enabled = true;
                    // textBox3.Enabled = true;

                    double tm = Convert.ToDouble(textBox1.Text);

                    textBox3.Text = (Convert.ToDouble(textBox1.Text) - Convert.ToDouble(textBox2.Text)).ToString();

                    try
                    {
                        cmd.CommandText = "Insert into [" + label5.Text + "] ([SN], [Date], [Particulars], [VCh_Type], [Debit],[Comments]) Values('" + label18.Text + "','" + dateTimePicker1.Text + "','" + comboBox2.Text + "','" + comboBox1.Text + "', '" + textBox1.Text + "','" + textBox4.Text + "')";
                        cmd.Connection = con;
                        cmd.ExecuteNonQuery();
                        cmd.CommandText = "Insert into [" + label5.Text + "] ([SN], [Particulars], [VCh_Type], [Credit]) Values('" + label18.Text + "','" + label6.Text + "','" + comboBox1.Text + "', '" + textBox2.Text + "')";
                        cmd.Connection = con;
                        cmd.ExecuteNonQuery();
                        cmd.CommandText = "Insert into [" + label5.Text + "] ([SN], [Particulars], [VCh_Type], [Credit]) Values('" + label18.Text + "','" + label7.Text + "','" + comboBox1.Text + "', '" + textBox3.Text + "')";
                        cmd.Connection = con;
                        cmd.ExecuteNonQuery();


                        try
                        {
                            cmd.CommandText = "Insert into [" + comboBox2.Text + "] ([SN], [Date], [Particulars], [Amount]) Values('" + label18.Text + "','" + dateTimePicker1.Text + "','" + label6.Text + "','" + textBox2.Text + "')";
                            cmd.Connection = con;
                            cmd.ExecuteNonQuery();
                            cmd.CommandText = "Insert into [" + comboBox2.Text + "] ([SN], [Date], [Particulars], [Amount]) Values('" + label18.Text + "','" + dateTimePicker1.Text + "','" + label7.Text + "','" + textBox3.Text + "')";
                            cmd.Connection = con;
                            cmd.ExecuteNonQuery();
                        }
                        catch (Exception)
                        {
                            cmd.CommandText = "CREATE TABLE [" + comboBox2.Text + "] ("+createaccout+")";
                            cmd.Connection = con;
                            cmd.ExecuteNonQuery();
                            cmd.CommandText = "Insert into [" + comboBox2.Text + "] ([SN], [Date], [Particulars], [Amount]) Values('" + label18.Text + "','" + dateTimePicker1.Text + "','" + label6.Text + "','" + textBox2.Text + "')";
                            cmd.Connection = con;
                            cmd.ExecuteNonQuery();
                            cmd.CommandText = "Insert into [" + comboBox2.Text + "] ([SN], [Date], [Particulars], [Amount]) Values('" + label18.Text + "','" + dateTimePicker1.Text + "','" + label7.Text + "','" + textBox3.Text + "')";
                            cmd.Connection = con;
                            cmd.ExecuteNonQuery();

                            cmd.CommandText = "insert into [Ptlist] ([Particulars]) values ('" + comboBox2.Text + "')";
                            cmd.Connection = con;
                            cmd.ExecuteNonQuery();
                        }
                        try
                        {

                            cmd.CommandText = "Insert into [" + label6.Text + "] ([SN_], [Date_], [Particulars_], [Amount_]) Values('" + label18.Text + "','" + dateTimePicker1.Text + "','" + comboBox2.Text + "','" + textBox2.Text + "')";
                            cmd.Connection = con;
                            cmd.ExecuteNonQuery();
                        }
                        catch (Exception)
                        {
                            cmd.CommandText = "CREATE TABLE [" + label6.Text + "] ("+createaccout+")";
                            cmd.Connection = con;
                            cmd.ExecuteNonQuery();
                            cmd.CommandText = "Insert into [" + label6.Text + "] ([SN_], [Date_], [Particulars_], [Amount_]) Values('" + label18.Text + "','" + dateTimePicker1.Text + "','" + comboBox2.Text + "','" + textBox2.Text + "')";
                            cmd.Connection = con;
                            cmd.ExecuteNonQuery();

                            cmd.CommandText = "insert into [Ptlist] ([Particulars]) values ('" +label6.Text + "')";
                            cmd.Connection = con;
                            cmd.ExecuteNonQuery();
                        }
                        try
                        {

                            cmd.CommandText = "Insert into [" + label7.Text + "] ([SN_], [Date_], [Particulars_], [Amount_]) Values('" + label18.Text + "','" + dateTimePicker1.Text + "','" + comboBox2.Text + "','" + textBox3.Text + "')";
                            cmd.Connection = con;
                            cmd.ExecuteNonQuery();
                        }
                        catch (Exception)
                        {
                            cmd.CommandText = "CREATE TABLE [" + label7.Text + "] ("+createaccout+")";
                            cmd.Connection = con;
                            cmd.ExecuteNonQuery();
                            cmd.CommandText = "Insert into [" + label7.Text + "] ([SN_], [Date_], [Particulars_], [Amount_]) Values('" + label18.Text + "','" + dateTimePicker1.Text + "','" + comboBox2.Text + "','" + textBox3.Text + "')";
                            cmd.Connection = con;
                            cmd.ExecuteNonQuery();
                            cmd.CommandText = "insert into [Ptlist] ([Particulars]) values ('" +label7.Text + "')";
                            cmd.Connection = con;
                            cmd.ExecuteNonQuery();
                        }

                        con.Close();
                        greedview();
                        textBox1.Clear();
                        textBox2.Clear();
                        textBox3.Clear();
                        textBox4.Clear();
                        comboBox2.Text = "";
                    }
                    catch (Exception)
                    {
                        ///MessageBox.Show(ex.ToString());

                    }
                }
                else
                {
                    MessageBox.Show("Please Enter The Amount and \n cash amount");
                    textBox1.Clear();
                    textBox2.Clear();
                    textBox3.Clear();
                    textBox4.Clear();
                    comboBox2.Text = "";
                }

            }
            else
            {
                MessageBox.Show("Invalid Selection amount Type !");
                textBox1.Clear();
                textBox2.Clear();
                textBox3.Clear();
                textBox4.Clear();
                comboBox2.Text = "";
            }
            con.Close();
        }


        public void receipt()
        {

            OleDbConnection con = new OleDbConnection();
            con.ConnectionString = @"Provider = Microsoft.ACE.OLEDB.12.0; Data Source =db.accdb";
            OleDbCommand cmd = con.CreateCommand();
            con.Open();

            if (comboBox3.SelectedIndex == 0)// cash
            {
                textBox2.Enabled = false;
                textBox2.Text = textBox1.Text;
                try
                {
                    cmd.CommandText = "Insert into [" + label5.Text + "] ([SN], [Date], [Particulars], [VCh_Type], [Credit], [Comments]) Values('" + label18.Text + "','" + dateTimePicker1.Text + "','" + comboBox2.Text + "','" + comboBox1.Text + "','" + textBox1.Text + "','" + textBox4.Text + "')";
                    cmd.Connection = con;
                    cmd.ExecuteNonQuery();
                    cmd.CommandText = "Insert into [" + label5.Text + "] ([SN], [Particulars], [VCh_Type], [Debit]) Values('" + label18.Text + "','" + comboBox3.Text + "','" + comboBox1.Text + "','" + textBox2.Text + "')";
                    cmd.Connection = con;
                    cmd.ExecuteNonQuery();

                    try
                    {
                        cmd.CommandText = "Insert into [" + comboBox3.Text + "] ([SN], [Date], [Particulars], [Amount]) Values('" + label18.Text + "','" + dateTimePicker1.Text + "','" + comboBox2.Text + "','" + textBox1.Text + "')";
                        cmd.Connection = con;
                        cmd.ExecuteNonQuery();
                    }
                    catch (Exception)
                    {

                        cmd.CommandText = "CREATE TABLE [" + comboBox3.Text + "] ("+createaccout+")";
                        cmd.Connection = con;
                        cmd.ExecuteNonQuery();
                        cmd.CommandText = "Insert into [" + comboBox3.Text + "] ([SN], [Date], [Particulars], [Amount]) Values('" + label18.Text + "','" + dateTimePicker1.Text + "','" + comboBox2.Text + "','" + textBox1.Text + "')";
                        cmd.Connection = con;
                        cmd.ExecuteNonQuery();
                        cmd.CommandText = "insert into [Ptlist] ([Particulars]) values ('" + comboBox3.Text + "')";
                        cmd.Connection = con;
                        cmd.ExecuteNonQuery();
                    }
                    try
                    {

                        cmd.CommandText = "Insert into [" + comboBox2.Text + "] ([SN_], [Date_], [Particulars_], [Amount_]) Values('" + label18.Text + "','" + dateTimePicker1.Text + "','" + comboBox3.Text + "','" + textBox1.Text + "')";
                        cmd.Connection = con;
                        cmd.ExecuteNonQuery();
                    }
                    catch (Exception)
                    {
                        cmd.CommandText = "CREATE TABLE [" + comboBox2.Text + "] ("+createaccout+")";
                        cmd.Connection = con;
                        cmd.ExecuteNonQuery();
                        cmd.CommandText = "Insert into [" + comboBox2.Text + "] ([SN_], [Date_], [Particulars_], [Amount_]) Values('" + label18.Text + "','" + dateTimePicker1.Text + "','" + comboBox3.Text + "','" + textBox1.Text + "')";
                        cmd.Connection = con;
                        cmd.ExecuteNonQuery();
                        cmd.CommandText = "insert into [Ptlist] ([Particulars]) values ('" + comboBox2.Text + "')";
                        cmd.Connection = con;
                        cmd.ExecuteNonQuery();
                    }
                    con.Close();
                    greedview();

                    textBox1.Clear();
                    textBox2.Clear();
                    textBox3.Clear();
                    textBox4.Clear();
                    textBox2.Enabled = false;
                    comboBox2.Text = "";


                }
                catch (Exception)
                {
                    //MessageBox.Show(ex.ToString());


                    con.Close();
                    greedview();

                    textBox1.Clear();
                    textBox2.Clear();
                    textBox3.Clear();
                    textBox4.Clear();
                    textBox2.Enabled = false;
                    comboBox2.Text = "";

                }
            }

            else if (comboBox3.Text == "Bank") //bank
            {
                //textBox3.Text = textBox1.Text;
                textBox2.Enabled = false;
                try
                {
                    cmd.CommandText = "Insert into [" + label5.Text + "] ([SN], [Date], [Particulars], [VCh_Type], [Credit], [Comments]) Values('" + label18.Text + "','" + dateTimePicker1.Text + "','" + comboBox2.Text + "','" + comboBox1.Text + "', '" + textBox1.Text + "','" + textBox4.Text + "')";
                    cmd.Connection = con;
                    cmd.ExecuteNonQuery();
                    cmd.CommandText = "Insert into [" + label5.Text + "] ([SN], [Particulars], [VCh_Type], [Debit]) Values('" + label18.Text + "','" + comboBox3.Text + "','" + comboBox1.Text + "','" + textBox1.Text + "')";
                    cmd.Connection = con;
                    cmd.ExecuteNonQuery();


                    try
                    {
                        cmd.CommandText = "Insert into [" + comboBox3.Text + "] ([SN], [Date], [Particulars], [Amount]) Values('" + label18.Text + "','" + dateTimePicker1.Text + "','" + comboBox2.Text + "','" + textBox1.Text + "')";
                        cmd.Connection = con;
                        cmd.ExecuteNonQuery();
                    }
                    catch (Exception)
                    {

                        cmd.CommandText = "CREATE TABLE [" + comboBox3.Text + "] ("+createaccout+")";
                        cmd.Connection = con;
                        cmd.ExecuteNonQuery();
                        cmd.CommandText = "Insert into [" + comboBox3.Text + "] ([SN], [Date], [Particulars], [Amount]) Values('" + label18.Text + "','" + dateTimePicker1.Text + "','" + comboBox2.Text + "','" + textBox1.Text + "')";
                        cmd.Connection = con;
                        cmd.ExecuteNonQuery();
                        cmd.CommandText = "insert into [Ptlist] ([Particulars]) values ('" + comboBox3.Text + "')";
                        cmd.Connection = con;
                        cmd.ExecuteNonQuery();
                    }
                    try
                    {

                        cmd.CommandText = "Insert into [" + comboBox2.Text + "] ([SN_], [Date_], [Particulars_], [Amount_]) Values('" + label18.Text + "','" + dateTimePicker1.Text + "','" + comboBox3.Text + "','" + textBox1.Text + "')";
                        cmd.Connection = con;
                        cmd.ExecuteNonQuery();
                    }
                    catch (Exception)
                    {
                        cmd.CommandText = "CREATE TABLE [" + comboBox2.Text + "] ("+createaccout+")";
                        cmd.Connection = con;
                        cmd.ExecuteNonQuery();
                        cmd.CommandText = "Insert into [" + comboBox2.Text + "] ([SN_], [Date_], [Particulars_], [Amount_]) Values('" + label18.Text + "','" + dateTimePicker1.Text + "','" + comboBox3.Text + "','" + textBox1.Text + "')";
                        cmd.Connection = con;
                        cmd.ExecuteNonQuery();
                        cmd.CommandText = "insert into [Ptlist] ([Particulars]) values ('" + comboBox2.Text + "')";
                        cmd.Connection = con;
                        cmd.ExecuteNonQuery();
                    }
                    con.Close();
                    greedview();

                    textBox1.Clear();
                    textBox2.Clear();
                    textBox3.Clear();
                    textBox4.Clear();
                    textBox2.Enabled = false;
                    comboBox2.Text = "";
                }
                catch (Exception)
                {
                    //MessageBox.Show(ex.ToString());


                    con.Close();
                    greedview();
                    textBox1.Clear();
                    textBox2.Clear();
                    textBox3.Clear();
                    textBox4.Clear();
                    textBox2.Enabled = false;
                    comboBox2.Text = "";
                }
            }


            else if (comboBox3.SelectedIndex == 2)// cash and bank
            {
                if (textBox1.Text != "" && textBox2.Text != "")
                {
                    textBox2.Enabled = true;
                    // textBox3.Enabled = true;

                    double tm = Convert.ToDouble(textBox1.Text);

                    textBox3.Text = (Convert.ToDouble(textBox1.Text) - Convert.ToDouble(textBox2.Text)).ToString();

                    try
                    {
                        cmd.CommandText = "Insert into [" + label5.Text + "] ([SN], [Date], [Particulars], [VCh_Type], [Credit],[Comments]) Values('" + label18.Text + "','" + dateTimePicker1.Text + "','" + comboBox2.Text + "','" + comboBox1.Text + "', '" + textBox1.Text + "','" + textBox4.Text + "')";
                        cmd.Connection = con;
                        cmd.ExecuteNonQuery();
                        cmd.CommandText = "Insert into [" + label5.Text + "] ([SN], [Particulars], [VCh_Type], [Debit]) Values('" + label18.Text + "','" + label6.Text + "','" + comboBox1.Text + "', '" + textBox2.Text + "')";
                        cmd.Connection = con;
                        cmd.ExecuteNonQuery();
                        cmd.CommandText = "Insert into [" + label5.Text + "] ([SN], [Particulars], [VCh_Type], [Debit]) Values('" + label18.Text + "','" + label7.Text + "','" + comboBox1.Text + "', '" + textBox3.Text + "')";
                        cmd.Connection = con;
                        cmd.ExecuteNonQuery();


                        try
                        {
                            cmd.CommandText = "Insert into [" + comboBox2.Text + "] ([SN_], [Date_], [Particulars_], [Amount_]) Values('" + label18.Text + "','" + dateTimePicker1.Text + "','" + label6.Text + "','" + textBox2.Text + "')";
                            cmd.Connection = con;
                            cmd.ExecuteNonQuery();
                            cmd.CommandText = "Insert into [" + comboBox2.Text + "] ([SN_], [Date_], [Particulars_], [Amount_]) Values('" + label18.Text + "','" + dateTimePicker1.Text + "','" + label7.Text + "','" + textBox3.Text + "')";
                            cmd.Connection = con;
                            cmd.ExecuteNonQuery();
                        }
                        catch (Exception)
                        {
                            cmd.CommandText = "CREATE TABLE [" + comboBox2.Text + "] ("+createaccout+")";
                            cmd.Connection = con;
                            cmd.ExecuteNonQuery();
                            cmd.CommandText = "Insert into [" + comboBox2.Text + "] ([SN_], [Date_], [Particulars_], [Amount_]) Values('" + label18.Text + "','" + dateTimePicker1.Text + "','" + label6.Text + "','" + textBox2.Text + "')";
                            cmd.Connection = con;
                            cmd.ExecuteNonQuery();
                            cmd.CommandText = "Insert into [" +comboBox2.Text + "] ([SN_], [Date_], [Particulars_], [Amount_]) Values('" + label18.Text + "','" + dateTimePicker1.Text + "','" + label7.Text + "','" + textBox3.Text + "')";
                            cmd.Connection = con;
                            cmd.ExecuteNonQuery();

                            cmd.CommandText = "insert into [Ptlist] ([Particulars]) values ('" + comboBox2.Text + "')";
                            cmd.Connection = con;
                            cmd.ExecuteNonQuery();
                        }
                        try
                        {

                            cmd.CommandText = "Insert into [" + label6.Text + "] ([SN], [Date], [Particulars], [Amount]) Values('" + label18.Text + "','" + dateTimePicker1.Text + "','" + comboBox2.Text + "','" + textBox2.Text + "')";
                            cmd.Connection = con;
                            cmd.ExecuteNonQuery();
                        }
                        catch (Exception)
                        {
                            cmd.CommandText = "CREATE TABLE [" + label6.Text + "] ("+createaccout+")";
                            cmd.Connection = con;
                            cmd.ExecuteNonQuery();
                            cmd.CommandText = "Insert into [" + label6.Text + "] ([SN], [Date], [Particulars], [Amount]) Values('" + label18.Text + "','" + dateTimePicker1.Text + "','" + comboBox2.Text + "','" + textBox2.Text + "')";
                            cmd.Connection = con;
                            cmd.ExecuteNonQuery();

                            cmd.CommandText = "insert into [Ptlist] ([Particulars]) values ('" +label6.Text + "')";
                            cmd.Connection = con;
                            cmd.ExecuteNonQuery();
                        }
                        try
                        {

                            cmd.CommandText = "Insert into [" + label7.Text + "] ([SN], [Date], [Particulars], [Amount]) Values('" + label18.Text + "','" + dateTimePicker1.Text + "','" + comboBox2.Text + "','" + textBox3.Text + "')";
                            cmd.Connection = con;
                            cmd.ExecuteNonQuery();
                        }
                        catch (Exception)
                        {
                            cmd.CommandText = "CREATE TABLE [" + label7.Text + "] ("+createaccout+")";
                            cmd.Connection = con;
                            cmd.ExecuteNonQuery();
                            cmd.CommandText = "Insert into [" + label7.Text + "] ([SN], [Date], [Particulars], [Amount]) Values('" + label18.Text + "','" + dateTimePicker1.Text + "','" + comboBox2.Text + "','" + textBox3.Text + "')";
                            cmd.Connection = con;
                            cmd.ExecuteNonQuery();
                            cmd.CommandText = "insert into [Ptlist] ([Particulars]) values ('" +label7.Text + "')";
                            cmd.Connection = con;
                            cmd.ExecuteNonQuery();
                        }

                        con.Close();
                        greedview();
                        textBox1.Clear();
                        textBox2.Clear();
                        textBox3.Clear();
                        textBox4.Clear();
                        comboBox2.Text = "";
                    }
                    catch (Exception)
                    {
                        ///MessageBox.Show(ex.ToString());

                    }
                }
                else
                {
                    MessageBox.Show("Please Enter The Amount and \n cash amount");
                    textBox1.Clear();
                    textBox2.Clear();
                    textBox3.Clear();
                    textBox4.Clear();
                    comboBox2.Text = "";
                }

            }
            else
            {
                MessageBox.Show("Invalid Selection amount Type !");
                textBox1.Clear();
                textBox2.Clear();
                textBox3.Clear();
                textBox4.Clear();
                comboBox2.Text = "";
            }
            con.Close();
        }
        public void contra()
        {
            //textBox2.Enabled = false;
            OleDbConnection con = new OleDbConnection();
            con.ConnectionString = @"Provider = Microsoft.ACE.OLEDB.12.0; Data Source =db.accdb";
            OleDbCommand cmd = con.CreateCommand();
            con.Open();


            if ((comboBox2.Text == "Cash")||(comboBox2.Text=="cash"))
            {
                try
                {
                    cmd.CommandText = "Insert into [" + label5.Text + "] ([SN], [Date], [Particulars], [VCh_Type], [Debit], [Comments]) Values('" + label18.Text + "','" + dateTimePicker1.Text + "','" + comboBox2.Text + "','" + comboBox1.Text + "','" + textBox1.Text + "','" + textBox4.Text + "')";
                    cmd.Connection = con;
                    cmd.ExecuteNonQuery();
                    cmd.CommandText = "Insert into [" + label5.Text + "] ([SN], [Particulars], [VCh_Type], [Credit]) Values('" + label18.Text + "','" + label7.Text + "','" + comboBox1.Text + "','" + textBox1.Text + "')";
                    cmd.Connection = con;
                    cmd.ExecuteNonQuery();

                    try
                    {
                        cmd.CommandText = "Insert into [" + comboBox2.Text + "] ([SN], [Date], [Particulars], [Amount]) Values('" + label18.Text + "','" + dateTimePicker1.Text + "','" + label7.Text + "','" + textBox1.Text + "')";
                        cmd.Connection = con;
                        cmd.ExecuteNonQuery();
                    }
                    catch (Exception)
                    {

                        cmd.CommandText = "CREATE TABLE [" +comboBox2.Text + "] ("+createaccout+")";
                        cmd.Connection = con;
                        cmd.ExecuteNonQuery();
                        cmd.CommandText = "Insert into [" + comboBox2.Text + "] ([SN], [Date], [Particulars], [Amount]) Values('" + label18.Text + "','" + dateTimePicker1.Text + "','" + label7.Text + "','" + textBox1.Text + "')";
                        cmd.Connection = con;
                        cmd.ExecuteNonQuery();
                        cmd.CommandText = "insert into [Ptlist] ([Particulars]) values ('" + comboBox2.Text + "')";
                        cmd.Connection = con;
                        cmd.ExecuteNonQuery();
                    }
                    try
                    {

                        cmd.CommandText = "Insert into [" + label7.Text + "] ([SN_], [Date_], [Particulars_], [Amount_]) Values('" + label18.Text + "','" + dateTimePicker1.Text + "','" + comboBox2.Text + "','" + textBox1.Text + "')";
                        cmd.Connection = con;
                        cmd.ExecuteNonQuery();
                    }
                    catch (Exception)
                    {
                        cmd.CommandText = "CREATE TABLE [" + label7.Text + "] ("+createaccout+")";
                        cmd.ExecuteNonQuery();
                        cmd.CommandText = "Insert into [" + label7.Text + "] ([SN_], [Date_], [Particulars_], [Amount_]) Values('" + label18.Text + "','" + dateTimePicker1.Text + "','" + comboBox2.Text + "','" + textBox1.Text + "')";
                        cmd.Connection = con;
                        cmd.ExecuteNonQuery();
                        cmd.CommandText = "insert into [Ptlist] ([Particulars]) values ('" +label7.Text + "')";
                        cmd.Connection = con;
                        cmd.ExecuteNonQuery();
                    }
                    con.Close();
                    greedview();

                    textBox1.Clear();
                    textBox2.Clear();
                    textBox3.Clear();
                    textBox4.Clear();
                    textBox2.Enabled = false;
                    comboBox2.Text = "";


                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());


                    con.Close();
                    greedview();

                    textBox1.Clear();
                    textBox2.Clear();
                    textBox3.Clear();
                    textBox4.Clear();
                    textBox2.Enabled = false;
                    comboBox2.Text = "";

                }

            }

            else if (comboBox2.Text == "Bank")
            {
                try
                {
                    cmd.CommandText = "Insert into [" + label5.Text + "] ([SN], [Date], [Particulars], [VCh_Type], [Debit], [Comments]) Values('" + label18.Text + "','" + dateTimePicker1.Text + "','" + comboBox2.Text + "','" + comboBox1.Text + "','" + textBox1.Text + "','" + textBox4.Text + "')";
                    cmd.Connection = con;
                    cmd.ExecuteNonQuery();
                    cmd.CommandText = "Insert into [" + label5.Text + "] ([SN], [Particulars], [VCh_Type], [Credit]) Values('" + label18.Text + "','" + label6.Text + "','" + comboBox1.Text + "','" + textBox1.Text + "')";
                    cmd.Connection = con;
                    cmd.ExecuteNonQuery();

                    try
                    {
                        cmd.CommandText = "Insert into [" +comboBox2.Text + "] ([SN], [Date], [Particulars], [Amount]) Values('" + label18.Text + "','" + dateTimePicker1.Text + "','" +label6.Text + "','" + textBox1.Text + "')";
                        cmd.Connection = con;
                        cmd.ExecuteNonQuery();
                    }
                    catch (Exception)
                    {

                        cmd.CommandText = "CREATE TABLE [" +comboBox2.Text + "] ("+createaccout+")";
                        cmd.Connection = con;
                        cmd.ExecuteNonQuery();
                        cmd.CommandText = "Insert into [" +comboBox2.Text + "] ([SN], [Date], [Particulars], [Amount]) Values('" + label18.Text + "','" + dateTimePicker1.Text + "','" + label6.Text + "','" + textBox1.Text + "')";
                        cmd.Connection = con;
                        cmd.ExecuteNonQuery();

                        cmd.CommandText = "insert into [Ptlist] ([Particulars]) values ('" + comboBox2.Text + "')";
                        cmd.Connection = con;
                        cmd.ExecuteNonQuery();
                    }
                    try
                    {

                        cmd.CommandText = "Insert into [" + label6.Text + "] ([SN_], [Date_], [Particulars_], [Amount_]) Values('" + label18.Text + "','" + dateTimePicker1.Text + "','" +comboBox2.Text + "','" + textBox1.Text + "')";
                        cmd.Connection = con;
                        cmd.ExecuteNonQuery();
                    }
                    catch (Exception)
                    {
                        cmd.CommandText = "CREATE TABLE [" +  label6.Text + "] ("+createaccout+")";
                        cmd.ExecuteNonQuery();
                        cmd.CommandText = "Insert into [" + label6.Text + "] ([SN_], [Date_], [Particulars_], [Amount_]) Values('" + label18.Text + "','" + dateTimePicker1.Text + "','" + comboBox2.Text + "','" + textBox1.Text + "')";
                        cmd.Connection = con;
                        cmd.ExecuteNonQuery();

                        cmd.CommandText = "insert into [Ptlist] ([Particulars]) values ('" +label6.Text + "')";
                        cmd.Connection = con;
                        cmd.ExecuteNonQuery();
                    }
                    con.Close();
                    greedview();

                    textBox1.Clear();
                    textBox2.Clear();
                    textBox3.Clear();
                    textBox4.Clear();
                    textBox2.Enabled = false;
                    comboBox2.Text = "";


                }
                catch (Exception)
                {
                    //MessageBox.Show(ex.ToString());


                    con.Close();
                    greedview();

                    textBox1.Clear();
                    textBox2.Clear();
                    textBox3.Clear();
                    textBox4.Clear();
                    textBox2.Enabled = false;
                    comboBox2.Text = "";

                }

            }

            else
            {
                MessageBox.Show("Please select Cash or bank Only ");
                textBox1.Clear();
                textBox2.Clear();
                textBox3.Clear();
                textBox4.Clear();
                textBox2.Enabled = false;
                comboBox2.Text = "";
            }
         
}
        public void sn()
        {

            OleDbConnection con = new OleDbConnection();
            con.ConnectionString = @"Provider = Microsoft.ACE.OLEDB.12.0; Data Source =db.accdb";

            OleDbCommand cmmd = new OleDbCommand("Select Max(ID) from [" + label5.Text + "]", con);
            con.Open();
            OleDbDataReader dr = cmmd.ExecuteReader();
            if (dr.Read())
            {
                label18.Text = dr[0].ToString();
                if (label18.Text == "")
                {
                    label18.Text = label5.Text + "-" +"1";
                }
                else
                {
                    try
            {



                label18.Text =label5.Text + "-" +(Convert.ToDouble(label18.Text)+1).ToString();
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

                }
                con.Close();

            }

        }
        
        private void button1_Click(object sender, EventArgs e)
        {

            sn();

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
                    while(comboBox2.Text != dt.Rows[i].ItemArray[1].ToString())
                    {
                      
                        i++;
                    }
                }
            }

            catch (Exception)
            {
                MessageBox.Show("This Particular is not In the list ! \n Please add particular ");
                textBox1.Clear();
                textBox2.Clear();
                textBox3.Clear();
                textBox4.Clear();
                textBox2.Enabled = false;
                comboBox2.Text = "";

            }



            OleDbCommand cmd = con.CreateCommand();
            // con.Open();

            if (textBox1.Text != "" && comboBox2.Text != "")
            {

                if (comboBox1.Text=="Payment") //pament

                {
                    comboBox3.Enabled = true;
                    payment();


                }

                else if (comboBox1.Text=="Receipt") //receipt
                {
                    comboBox3.Enabled = true;
                    receipt();

                }

                else if (comboBox1.Text=="Contra") //contra

                {
                    contra();
                }
                else
                {
                    MessageBox.Show("Invalid vauchar type selection\n Please select Payment, Receipt or Contra ");

                    textBox1.Clear();
                    textBox2.Clear();
                    textBox3.Clear();
                    textBox4.Clear();
                    comboBox2.Text = "";
                }
            }
            else
            {
                MessageBox.Show("Please select the particulars from the list and \nPlease enter the Amount");

                textBox1.Clear();
                textBox2.Clear();
                textBox3.Clear();
                textBox4.Clear();
                comboBox2.Text = "";
            }


            dtotal();
            ctotal();
            con.Close();
            comboBox3.SelectedIndex = 0;

        }

        private void comboBox2_Click(object sender, EventArgs e)
        {
            comboBox2.DroppedDown = true;
        }

        private void comboBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            comboBox2.DroppedDown = true;
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
       
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

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            if(comboBox3.SelectedIndex==2)
            {
                textBox2.Enabled = true;
            }
            else
            {
                textBox2.Enabled = false;
            }
        }

        private void comboBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            comboBox1.DroppedDown = true;
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            if (textBox1.Text != "")
            {
                if (textBox2.Text != "")
                {
                    textBox3.Text = (Convert.ToDouble(textBox1.Text) - Convert.ToDouble(textBox2.Text)).ToString();
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Form3.tbn = label5.Text;
            frmVaucherMd vf = new frmVaucherMd();
            vf.Show();
            this.Hide();

        }

        private void button3_Click(object sender, EventArgs e)
        {
            frmParticulars adpt = new frmParticulars();
            adpt.Show();
            this.Hide();
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Form2 frm2 = new Form2();
            frm2.Show();
            this.Hide();
        }
        Int32 itemperpage = 0;
        Int32 totalrow = 0;
       Int32 rows = 0;
        private void button4_Click(object sender, EventArgs e)
        {
            greedview();
            //totalrow = dataGridView1.Rows.Count;
            itemperpage = rows = 0;
            printPreviewDialog1.Document = printDocument1;

           // ((ToolStripButton)((ToolStrip)printPreviewDialog1.Controls[1]).Items[0]).Enabled
             // = false;

            printPreviewDialog1.ShowDialog();

        }

     

        private void printDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            greedview();
            //int a = dataGridView1.Rows.Count;




            e.Graphics.DrawString("Assemblies of God Of Nepal", new Font("Ariel", 12, FontStyle.Bold), Brushes.Black, new Point(310, 30));
            e.Graphics.DrawString("State No.7", new Font("Ariel", 14, FontStyle.Bold), Brushes.Black, new Point(370, 50));
            e.Graphics.DrawString("Farwestern Regional AGN", new Font("Ariel", 28, FontStyle.Bold), Brushes.Black, new Point(170, 66));
            e.Graphics.DrawString("Dhangadhi, Kailali, Nepal", new Font("Ariel", 18, FontStyle.Bold), Brushes.Black, new Point(280, 110));
            e.Graphics.DrawString("Vauchar", new Font("Ariel", 24, FontStyle.Bold), Brushes.Black, new Point(350, 140));
            e.Graphics.DrawString("Month of :", new Font("Ariel", 16, FontStyle.Bold), Brushes.Black, new Point(30, 200));
            e.Graphics.DrawString(label5.Text, new Font("Ariel", 18, FontStyle.Bold), Brushes.Black, new Point(370, 190));
            e.Graphics.DrawString("-------------------------------------------------------------------------------------------------------------------------------------------"
                , new Font("Ariel", 12, FontStyle.Regular), Brushes.Black, new Point(25, 235));

            e.Graphics.DrawString("Date", new Font("Ariel", 13, FontStyle.Bold), Brushes.Black, new Point(30, 255));
            e.Graphics.DrawString("Particulars", new Font("Ariel", 13, FontStyle.Bold), Brushes.Black, new Point(130, 255));
            e.Graphics.DrawString("Vch_Type", new Font("Ariel", 13, FontStyle.Bold), Brushes.Black, new Point(350, 255));
            e.Graphics.DrawString("JF", new Font("Ariel", 13, FontStyle.Bold), Brushes.Black, new Point(450, 255));
            e.Graphics.DrawString("Debit", new Font("Ariel", 13, FontStyle.Bold), Brushes.Black, new Point(480, 255));
            e.Graphics.DrawString("Credit", new Font("Ariel", 13, FontStyle.Bold), Brushes.Black, new Point(580, 255));
            e.Graphics.DrawString("Comments", new Font("Ariel", 13, FontStyle.Bold), Brushes.Black, new Point(680, 255));
            e.Graphics.DrawString("-------------------------------------------------------------------------------------------------------------------------------------------"
                , new Font("Ariel", 12, FontStyle.Regular), Brushes.Black, new Point(25, 270));

            int y = 300;

            string strProvider = @"Provider = Microsoft.ACE.OLEDB.12.0; Data Source = db.accdb";
            string strSql = "Select * from [" + label5.Text + "] order by ID asc";
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

            //for (; rows < a; rows++)
            while(rows<dataGridView1.Rows.Count)
            {

                try {

                    e.Graphics.DrawString(dataGridView1.Rows[rows].Cells[2].Value.ToString().Trim(), new Font("Ariel", 13, FontStyle.Bold), Brushes.Black, new Point(30, y));
                    e.Graphics.DrawString(dataGridView1.Rows[rows].Cells[3].Value.ToString().Trim(), new Font("Ariel", 13, FontStyle.Bold), Brushes.Black, new Point(150, y));
                    e.Graphics.DrawString(dataGridView1.Rows[rows].Cells[4].Value.ToString().Trim(), new Font("Ariel", 13, FontStyle.Bold), Brushes.Black, new Point(350, y));
                    e.Graphics.DrawString(dataGridView1.Rows[rows].Cells[5].Value.ToString().Trim(), new Font("Ariel", 13, FontStyle.Bold), Brushes.Black, new Point(450, y));
                    e.Graphics.DrawString(dataGridView1.Rows[rows].Cells[6].Value.ToString().Trim(), new Font("Ariel", 13, FontStyle.Bold), Brushes.Black, new Point(480, y));
                    e.Graphics.DrawString(dataGridView1.Rows[rows].Cells[7].Value.ToString().Trim(), new Font("Ariel", 13, FontStyle.Bold), Brushes.Black, new Point(580, y));
                    y += 30;
                    //rows++;
                    rows += 1;
                    if (itemperpage < 24)
                    {
                        itemperpage += 1;
                        e.HasMorePages = false;
                    }
                    else
                    {
                        itemperpage = 0;
                        e.HasMorePages = true;
                        return;
                    }

                }
                catch(Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
                
            }
            e.Graphics.DrawString("Total", new Font("Ariel", 13, FontStyle.Bold), Brushes.Black, new Point(150, y+5));
            e.Graphics.DrawString(label14.Text, new Font("Ariel", 13, FontStyle.Bold), Brushes.Black, new Point(480, y+5));
            e.Graphics.DrawString(label15.Text, new Font("Ariel", 13, FontStyle.Bold), Brushes.Black, new Point(580, y+5));

            greedview();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            itemperpage = rows = 0;
            //printDialog1.Document = printDocument1;
            //printDialog1.ShowDialog();
            printDialog1.Document = printDocument1;

            if (printDialog1.ShowDialog() == DialogResult.OK)
            {
                printDocument1.Print();
            }
            
        }

        private void button6_Click(object sender, EventArgs e)
        {
            frmDelete d = new frmDelete();
            d.Show();
            this.Hide();
        }
    }
}
