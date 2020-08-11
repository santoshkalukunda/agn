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
    public partial class Form4 : Form 
    {
       
        public Form4()
        {
            InitializeComponent();
        }

        private void Form4_Load(object sender, EventArgs e)
        {
            ptlist();
            greedview();
            button2.Hide();
            dtotal();
            ctotal();
            btotal();
            total();


            ebalancetotal();
            ibalancetotal();
            button4.Hide();
            button3.Hide();
            label1.Hide();
            label2.Hide();
            label3.Hide();
            label3.Hide();


        }
        public double summ;
        public void btotal()
        {
            
            OleDbConnection con = new OleDbConnection();
            con.ConnectionString = @"Provider = Microsoft.ACE.OLEDB.12.0; Data Source =db.accdb";
            OleDbCommand cmd = con.CreateCommand();
            con.Open();

             double sum = 0;
            DataTable dataTable = (DataTable)dataGridView1.DataSource;

try {
                DataRow drToAdd = dataTable.NewRow();


                if (Convert.ToDouble(label2.Text) < Convert.ToDouble(label1.Text))
                {

                    sum = Convert.ToDouble(label1.Text) - Convert.ToDouble(label2.Text);

                    drToAdd[7] = "Balance";
                    drToAdd[8] = sum;
                    summ = sum;


                    dataTable.Rows.Add(drToAdd);
                    dataTable.AcceptChanges();

                    expenditure();

                }
                else if(Convert.ToDouble(label2.Text) > Convert.ToDouble(label1.Text))
                {
                    sum = Convert.ToDouble(label2.Text) - Convert.ToDouble(label1.Text);
                    drToAdd[3] = "Balance";
                    drToAdd[4] = sum;
                    summ = sum;
                    dataTable.Rows.Add(drToAdd);
                    dataTable.AcceptChanges();


                    income();
                }
                else
                {

                }
           }
            catch
            {

                MessageBox.Show("The account table is '"+comboBox1.Text+"' not found !");

                OleDbConnection conn = new OleDbConnection();
                conn.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=db.accdb";
                conn.Open();
                string sqlstr = "Drop table [" +comboBox1.Text + "]";
                OleDbCommand cmmd = new OleDbCommand(sqlstr, conn);
                // cmmd.CommandText ;

                try
                {
                   
                    {
                        cmd.CommandText = "Delete from [Ptlist] where [Particulars]='" +comboBox1.Text + "'";
                        cmd.Connection = con;
                        cmd.ExecuteNonQuery();

                    }
                }
                catch
                {

                }
                if (conn.State == ConnectionState.Open)
                {
                    try
                    {

                        
                        {
                            cmmd.ExecuteNonQuery();


                        }
                        conn.Close();
                    }
                    catch (OleDbException )
                    {

                        conn.Close();
                    }
                }
                else
                {
                    MessageBox.Show("Error!");
                }
                try
                {
                    comboBox1.SelectedIndex++;
                }
                catch
                {

                }
                
            }

        }
        public void total()
        {
            OleDbConnection con = new OleDbConnection();
            con.ConnectionString = @"Provider = Microsoft.ACE.OLEDB.12.0; Data Source =db.accdb";
            OleDbCommand cmd = con.CreateCommand();
            con.Open();
            if (label1.Text == "0" && label2.Text == "0")
            {

                OleDbConnection conn = new OleDbConnection();
                conn.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=db.accdb";
                conn.Open();
                string sqlstr = "Drop table [" + comboBox1.Text + "]";
                OleDbCommand cmmd = new OleDbCommand(sqlstr, conn);
                // cmmd.CommandText ;

                try
                {

                    {
                        cmd.CommandText = "Delete from [Ptlist] where [Particulars]='" + comboBox1.Text + "'";
                        cmd.Connection = con;
                        cmd.ExecuteNonQuery();

                    }
                }
                catch
                {

                }
                if (conn.State == ConnectionState.Open)
                {
                    try
                    {


                        {
                            cmmd.ExecuteNonQuery();


                        }
                        conn.Close();
                    }
                    catch (OleDbException )
                    {

                        conn.Close();
                    }
                }
                else
                {
                    MessageBox.Show("Error!");
                }
                try
                {
                    comboBox1.SelectedIndex++;
                }
                catch
                {

                }
            }
            else { 
                try
                {
                    DataTable dataTable = (DataTable)dataGridView1.DataSource;
                    DataRow drToAdd = dataTable.NewRow();
                    if (Convert.ToDouble(label2.Text) < Convert.ToDouble(label1.Text))
                    {
                        drToAdd[4] = label1.Text;
                        drToAdd[3] = "Total";
                        drToAdd[7] = "Total";
                        drToAdd[8] = label1.Text;

                        dataTable.Rows.Add(drToAdd);
                        dataTable.AcceptChanges();

                    }
                    else
                    {
                        drToAdd[4] = label2.Text;
                        drToAdd[3] = "Total";
                        drToAdd[7] = "Total";
                        drToAdd[8] = label2.Text;

                        dataTable.Rows.Add(drToAdd);
                        dataTable.AcceptChanges();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
            }

        }
        public void income()
        {
            DataGridViewRow row = (DataGridViewRow)dataGridView2.Rows[0].Clone();
            row.Cells[0].Value = comboBox1.Text;
            row.Cells[1].Value = summ;
            dataGridView2.Rows.Add(row);
        }
        public void expenditure()
        {
            DataGridViewRow row = (DataGridViewRow)dataGridView3.Rows[0].Clone();
            row.Cells[0].Value = comboBox1.Text;
            row.Cells[1].Value = summ;
            dataGridView3.Rows.Add(row);

        }
      
        public void greedview()
        {

            string strProvider = @"Provider = Microsoft.ACE.OLEDB.12.0; Data Source = db.accdb";
            string strSql = "Select * from ["+comboBox1.Text+"]";
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
        public void dtotal()
        {
            double sum = 0;
            for (int i = 0; i < dataGridView1.Rows.Count; ++i)
            {
                if (Convert.ToString(dataGridView1.Rows[i].Cells[4].Value) == "")
                {
                    continue;
                }
                else
                {
                    sum += Convert.ToDouble(dataGridView1.Rows[i].Cells[4].Value);
                }
            }
            label1.Text = sum.ToString();

        }
        public void ctotal()
        {
            double sum = 0;
            for (int i = 0; i < dataGridView1.Rows.Count; ++i)
            {
                if (Convert.ToString(dataGridView1.Rows[i].Cells[8].Value) == "")
                {
                    continue;
                }
                else
                {
                    sum += Convert.ToDouble(dataGridView1.Rows[i].Cells[8].Value);
                }
            }
            label2.Text = sum.ToString();

        }
        public void ebalancetotal()
        {
            double sum = 0;
            for(int i=0; i<dataGridView2.Rows.Count; i++)
            {
                if (Convert.ToString(dataGridView2.Rows[i].Cells[1].Value) == "")
                {
                    continue;
                }
                else
                {
                    sum += Convert.ToDouble(dataGridView2.Rows[i].Cells[1].Value);
                }
            }
            label3.Text = sum.ToString();
        }
        public void ibalancetotal()
        {
            double sum = 0;
            for (int i = 0; i < dataGridView3.Rows.Count; i++)
            {
                if (Convert.ToString(dataGridView3.Rows[i].Cells[1].Value) == "")
                {
                    continue;
                }
                else
                {
                    sum += Convert.ToDouble(dataGridView3.Rows[i].Cells[1].Value);
                }
            }
            label4.Text = sum.ToString();
        }


        public void ptlist()
        {
            OleDbConnection con = new OleDbConnection();
            string q = "select Particulars from Ptlist";
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
                    //comboBox1.Text = "";
                    con.Close();

                }
            }
            catch (Exception)
            {
                // MessageBox.Show(ex.ToString());
            }

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
        int i = 1;
        private void button1_Click(object sender, EventArgs e)
        {
            button2.Hide();
            //int i = 1;
            try
            {
             
                
                    comboBox1.SelectedIndex = i;
                    i++;

                    greedview();


                    dtotal();
                    ctotal();
                    btotal();
                    total();
                    ibalancetotal();
                    ebalancetotal();
                
            }
            catch(Exception)
            {
                MessageBox.Show("Finish account table");
                //i = 0;
                balance();
                button1.Hide();
                button2.Hide();
                button3.Show();
                button4.Show();
            }
            

        }

        public void balance()
        {
           
            
            if(Convert.ToDouble(label3.Text)> Convert.ToDouble(label4.Text))
            {
               
                DataGridViewRow row = (DataGridViewRow)dataGridView3.Rows[0].Clone();
                row.Cells[0].Value = "Deficit";
                row.Cells[1].Value = Convert.ToDouble(label3.Text) - Convert.ToDouble(label4.Text); ;
                dataGridView3.Rows.Add(row);

                DataGridViewRow row2 = (DataGridViewRow)dataGridView3.Rows[0].Clone();
                row2.Cells[0].Value = "Total";
                row2.Cells[1].Value = label3.Text;
                dataGridView3.Rows.Add(row2);


                DataGridViewRow row1 = (DataGridViewRow)dataGridView2.Rows[0].Clone();
                row1.Cells[0].Value = "Total";
                row1.Cells[1].Value = label3.Text;
                dataGridView2.Rows.Add(row1);


            }
            else if (Convert.ToDouble(label3.Text) < Convert.ToDouble(label4.Text))
            {
                DataGridViewRow row = (DataGridViewRow)dataGridView2.Rows[0].Clone();
                row.Cells[0].Value = "Surplus";
                row.Cells[1].Value = Convert.ToDouble(label4.Text) - Convert.ToDouble(label3.Text); ;
                dataGridView2.Rows.Add(row);

                DataGridViewRow row2 = (DataGridViewRow)dataGridView2.Rows[0].Clone();
                row2.Cells[0].Value = "Total";
                row2.Cells[1].Value = label4.Text;
                dataGridView2.Rows.Add(row2);

                DataGridViewRow row1 = (DataGridViewRow)dataGridView3.Rows[0].Clone();
                row1.Cells[0].Value = "Total";
                row1.Cells[1].Value = label4.Text;
                dataGridView3.Rows.Add(row1);
            }
            else
            {
                DataGridViewRow row = (DataGridViewRow)dataGridView2.Rows[0].Clone();
                row.Cells[0].Value = "Total";
                row.Cells[1].Value = label3.Text ;
                dataGridView2.Rows.Add(row);

                DataGridViewRow row1 = (DataGridViewRow)dataGridView3.Rows[0].Clone();
                row1.Cells[0].Value = "Total";
                row1.Cells[1].Value = label3.Text;
                dataGridView3.Rows.Add(row1);
            }
        }

        private void comboBox1_SelectionChangeCommitted(object sender, EventArgs e)
        {
            greedview();
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            //greedview();
        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            
            int i = 1;
            try
            {
                while (i < dataGridView1.Rows.Count)
                {
                    comboBox1.SelectedIndex = i;
                    i++;

                    greedview();


                    dtotal();
                    ctotal();
                    btotal();
                    total();
                    ibalancetotal();
                    ebalancetotal();
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Finish account table");
                //i = 0;
                balance();
                button1.Hide();
                button2.Hide();
                button3.Show();
                button4.Show();
            }

            
        }
        Int32 itemperpage = 0;
        Int32 page = 1;
        Int32 rows = 0;
        Int32 rows1 = 0;
        private void printDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {

            //totalrow = dataGridView2.Rows.Count;
            e.Graphics.DrawString("Assemblies of God Of Nepal", new Font("Ariel", 12, FontStyle.Bold), Brushes.Black, new Point(310, 30));
            e.Graphics.DrawString("State No.7", new Font("Ariel", 14, FontStyle.Bold), Brushes.Black, new Point(370, 50));
            e.Graphics.DrawString("Farwestern Regional AGN", new Font("Ariel", 28, FontStyle.Bold), Brushes.Black, new Point(170, 66));
            e.Graphics.DrawString("Dhangadhi, Kailali, Nepal", new Font("Ariel", 18, FontStyle.Bold), Brushes.Black, new Point(280, 110));
            e.Graphics.DrawString("Balance Sheet", new Font("Ariel", 24, FontStyle.Bold), Brushes.Black, new Point(320, 140));
            //e.Graphics.DrawString("Date:", new Font("Ariel", 16, FontStyle.Bold), Brushes.Black, new Point(30, 200));
            //e.Graphics.DrawString("B", new Font("Ariel", 18, FontStyle.Bold), Brushes.Black, new Point(370, 190));
            e.Graphics.DrawString("-------------------------------------------------------------------------------------------------------------------------------------------"
                , new Font("Ariel", 12, FontStyle.Regular), Brushes.Black, new Point(25, 235));

            e.Graphics.DrawString("Income", new Font("Ariel", 13, FontStyle.Bold), Brushes.Black, new Point(30, 255));
            e.Graphics.DrawString("Amount", new Font("Ariel", 13, FontStyle.Bold), Brushes.Black, new Point(250, 255));
            e.Graphics.DrawString("Expenditure", new Font("Ariel", 13, FontStyle.Bold), Brushes.Black, new Point(450, 255));
            e.Graphics.DrawString("Amount", new Font("Ariel", 13, FontStyle.Bold), Brushes.Black, new Point(650, 255));
            
            e.Graphics.DrawString("-------------------------------------------------------------------------------------------------------------------------------------------"
                , new Font("Ariel", 12, FontStyle.Regular), Brushes.Black, new Point(25, 270));
            int y = 300;

            while (rows < dataGridView2.Rows.Count-1 || rows1<dataGridView3.Rows.Count-1)
            {

                try
                {
                    if (rows < dataGridView2.Rows.Count - 2)
                    {

                        e.Graphics.DrawString(dataGridView2.Rows[rows].Cells[0].Value.ToString().Trim(), new Font("Ariel", 13, FontStyle.Bold), Brushes.Black, new Point(30, y));
                        e.Graphics.DrawString(dataGridView2.Rows[rows].Cells[1].Value.ToString().Trim(), new Font("Ariel", 13, FontStyle.Bold), Brushes.Black, new Point(240, y));
                    }
                    if (rows1 < dataGridView3.Rows.Count-2)
                    {
                        e.Graphics.DrawString(dataGridView3.Rows[rows1].Cells[0].Value.ToString().Trim(), new Font("Ariel", 13, FontStyle.Bold), Brushes.Black, new Point(450, y));
                        e.Graphics.DrawString(dataGridView3.Rows[rows1].Cells[1].Value.ToString().Trim(), new Font("Ariel", 13, FontStyle.Bold), Brushes.Black, new Point(650, y));
                    }
                        //e.Graphics.DrawString(dataGridView2.Rows[rows].Cells[6].Value.ToString().Trim(), new Font("Ariel", 13, FontStyle.Bold), Brushes.Black, new Point(480, y));
                    //e.Graphics.DrawString(dataGridView2.Rows[rows].Cells[7].Value.ToString().Trim(), new Font("Ariel", 13, FontStyle.Bold), Brushes.Black, new Point(580, y));
                    y += 30;
                    //rows++;
                    rows += 1;
                    rows1 += 1;
                    if (itemperpage < 24)
                    {
                        itemperpage += 1;
                        e.HasMorePages = false;
                    }
                    else
                    {
                        itemperpage = 0;
                        page += 1;
                        e.HasMorePages = true;
                        return;
                    }

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }

            }
            if (Convert.ToDouble(label3.Text) > Convert.ToDouble(label4.Text))
            {
                e.Graphics.DrawString("Total", new Font("Ariel", 13, FontStyle.Bold), Brushes.Black, new Point(30, y));
                e.Graphics.DrawString(label3.Text, new Font("Ariel", 13, FontStyle.Bold), Brushes.Black, new Point(250, y));
                e.Graphics.DrawString("Total", new Font("Ariel", 13, FontStyle.Bold), Brushes.Black, new Point(450, y));
                e.Graphics.DrawString(label4.Text, new Font("Ariel", 13, FontStyle.Bold), Brushes.Black, new Point(650, y));
            }
            else
            {
                e.Graphics.DrawString("Total", new Font("Ariel", 13, FontStyle.Bold), Brushes.Black, new Point(30, y));
                e.Graphics.DrawString(label4.Text, new Font("Ariel", 13, FontStyle.Bold), Brushes.Black, new Point(250, y));
                e.Graphics.DrawString("Total", new Font("Ariel", 13, FontStyle.Bold), Brushes.Black, new Point(450, y));
                e.Graphics.DrawString(label4.Text, new Font("Ariel", 13, FontStyle.Bold), Brushes.Black, new Point(650, y));
            }
            e.Graphics.DrawString(page.ToString(), new Font("Ariel", 13, FontStyle.Bold), Brushes.Black, new Point(500, 900));
        }

        private void button3_Click(object sender, EventArgs e)
        {
            itemperpage = rows=rows1= 0;
            printPreviewDialog1.Document = printDocument1;
            printPreviewDialog1.ShowDialog();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            itemperpage =rows= 0;
            //printDialog1.Document = printDocument1;
            //printDialog1.ShowDialog();
            printDialog1.Document = printDocument1;

            if (printDialog1.ShowDialog() == DialogResult.OK)
            {
                printDocument1.Print();
            }
        }

        private void printDocument2_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            e.Graphics.DrawString("Assemblies of God Of Nepal", new Font("Ariel", 12, FontStyle.Bold), Brushes.Black, new Point(310, 30));
            e.Graphics.DrawString("State No.7", new Font("Ariel", 14, FontStyle.Bold), Brushes.Black, new Point(370, 50));
            e.Graphics.DrawString("Farwestern Regional AGN", new Font("Ariel", 28, FontStyle.Bold), Brushes.Black, new Point(170, 66));
            e.Graphics.DrawString("Dhangadhi, Kailali, Nepal", new Font("Ariel", 18, FontStyle.Bold), Brushes.Black, new Point(280, 110));
          
            e.Graphics.DrawString("Account Table", new Font("Ariel", 24, FontStyle.Bold), Brushes.Black, new Point(300, 140));
            e.Graphics.DrawString(comboBox1.Text, new Font("Ariel", 16, FontStyle.Bold), Brushes.Black, new Point(380, 200));
            e.Graphics.DrawString("Dr.", new Font("Ariel", 13, FontStyle.Bold), Brushes.Black, new Point(30, 220));
            e.Graphics.DrawString("Cr.", new Font("Ariel", 13, FontStyle.Bold), Brushes.Black, new Point(720, 220));
            e.Graphics.DrawString("-------------------------------------------------------------------------------------------------------------------------------------------"
                , new Font("Ariel", 12, FontStyle.Regular), Brushes.Black, new Point(25, 235));

            e.Graphics.DrawString("Date", new Font("Ariel", 13, FontStyle.Bold), Brushes.Black, new Point(30, 255));
            e.Graphics.DrawString("Particulars", new Font("Ariel", 13, FontStyle.Bold), Brushes.Black, new Point(150, 255));
            e.Graphics.DrawString("Amount", new Font("Ariel", 13, FontStyle.Bold), Brushes.Black, new Point(300, 255));
            e.Graphics.DrawString("Date", new Font("Ariel", 13, FontStyle.Bold), Brushes.Black, new Point(450, 255));
            e.Graphics.DrawString("Particulars", new Font("Ariel", 13, FontStyle.Bold), Brushes.Black, new Point(560, 255));
            e.Graphics.DrawString("Amount", new Font("Ariel", 13, FontStyle.Bold), Brushes.Black, new Point(710, 255));

            e.Graphics.DrawString("-------------------------------------------------------------------------------------------------------------------------------------------"
                , new Font("Ariel", 12, FontStyle.Regular), Brushes.Black, new Point(25, 270));
            int y = 300;

            while (rows < dataGridView1.Rows.Count - 1)
            {

                try
                {
                    if (rows < dataGridView1.Rows.Count-1)
                    {

                        e.Graphics.DrawString(dataGridView1.Rows[rows].Cells[2].Value.ToString().Trim(), new Font("Ariel", 13, FontStyle.Bold), Brushes.Black, new Point(30, y));
                        e.Graphics.DrawString(dataGridView1.Rows[rows].Cells[3].Value.ToString().Trim(), new Font("Ariel", 13, FontStyle.Bold), Brushes.Black, new Point(150, y));
                  
                        e.Graphics.DrawString(dataGridView1.Rows[rows].Cells[4].Value.ToString().Trim(), new Font("Ariel", 13, FontStyle.Bold), Brushes.Black, new Point(300, y));
                        e.Graphics.DrawString(dataGridView1.Rows[rows].Cells[6].Value.ToString().Trim(), new Font("Ariel", 13, FontStyle.Bold), Brushes.Black, new Point(450, y));
                        e.Graphics.DrawString(dataGridView1.Rows[rows].Cells[7].Value.ToString().Trim(), new Font("Ariel", 13, FontStyle.Bold), Brushes.Black, new Point(560, y));
                        e.Graphics.DrawString(dataGridView1.Rows[rows].Cells[8].Value.ToString().Trim(), new Font("Ariel", 13, FontStyle.Bold), Brushes.Black, new Point(710, y));

                    }
                    y += 30;
                    rows++;
                   // rows += 1;
                   
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
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }

            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            itemperpage = rows = 0;
            printPreviewDialog1.Document = printDocument2;
            printPreviewDialog1.ShowDialog();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            itemperpage = rows = 0;
            
            printDialog1.Document = printDocument2;

            if (printDialog1.ShowDialog() == DialogResult.OK)
            {
                printDocument2.Print();
            }
        }
    }
    
}
