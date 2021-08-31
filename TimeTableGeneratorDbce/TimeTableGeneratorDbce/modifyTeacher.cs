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
    public partial class modifyTeacher : Form
    {
        public modifyTeacher()
        {
            InitializeComponent();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                OleDbConnection con = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; Data Source='D:\\course.xlsx';Extended Properties=Excel 8.0 ;");

                con.Open();
                String sql = "select * from [sheet2$]";
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
                String sql = "UPDATE [sheet2$] SET [Name]='" + a + "',[Course]='" + b + "',[FullName]='" + c + "',[Phno]='" + d + "' WHERE [Name]='" + x + "' ";
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
    }
}
