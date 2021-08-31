using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TimeTableGeneratorDbce
{
    public partial class modifyCourse : Form
    {
        public modifyCourse()
        {
            InitializeComponent();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {


        }


        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                OleDbConnection con = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; Data Source='D:\\course.xlsx';Extended Properties=Excel 8.0 ;");

                con.Open();
                String sql = "select * from [sheet1$]";
                OleDbDataAdapter ad = new OleDbDataAdapter(sql, con);
                //OleDbCommand cmd = new OleDbCommand(sql, con);
                DataSet ds = new DataSet();
                ad.Fill(ds);
                dataGridView1.DataSource = ds.Tables[0];
                //cmd.ExecuteNonQuery();
                con.Close();

            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {

            String x = textBox1.Text;
            String a = textBox2.Text;
            String b = textBox3.Text;
            String c = textBox4.Text;
            String d = textBox5.Text;


            try
            {
                OleDbConnection con = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; Data Source='D:\\course.xlsx';Extended Properties=Excel 8.0 ;");
                // OleDbDataAdapter ad = new OleDbDataAdapter();
                con.Open();
                String sql = "UPDATE [sheet1$] SET [Name]='" + a + "',[Year]='" + b + "',[Semester]='" + c + "',[Credit]='" + d + "' WHERE [Name]='" + x + "' ";
                OleDbCommand cmd = new OleDbCommand(sql, con);
                cmd.ExecuteNonQuery();
                con.Close();

            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            MessageBox.Show("Data Modified Successfully !");
        }

        private void label7_Click(object sender, EventArgs e)
        {

        }

        private void label6_Click(object sender, EventArgs e)
        {

        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }
    }
}

