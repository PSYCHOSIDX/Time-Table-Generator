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

namespace TimeTableGeneratorDbce
{
    public partial class addCourse : Form
    {
        public addCourse()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
           
            String a = textBox1.Text;
            String b = textBox2.Text;
            String c = textBox3.Text;
            String d = textBox4.Text;

       

            try
                        {
                            OleDbConnection con=new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; Data Source='D:\\course.xlsx';Extended Properties=Excel 8.0 ;");
                OleDbDataAdapter ad = new OleDbDataAdapter();
                con.Open();
                String sql = "insert into [sheet1$]([Name],[Year],[Semester],[Credit]) values('"+a+"','"+b+"','"+c+"','"+d+"')";
                OleDbCommand cmd=new OleDbCommand(sql,con);
                cmd.ExecuteNonQuery();
                con.Close();
               
                        }

                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.ToString());
                        }
            MessageBox.Show("Data Entered Successfully !");
        }

        private void label7_Click(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
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

        private void textBox4_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void label8_Click(object sender, EventArgs e)
        {

        }
    }
}
