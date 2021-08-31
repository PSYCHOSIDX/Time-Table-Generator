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
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace TimeTableGeneratorDbce
{
    public partial class courseDelete : Form
    {
        public courseDelete()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
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

        private void button2_Click(object sender, EventArgs e)
{
            String x = textBox1.Text;
            try
            {
                OleDbConnection con = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; Data Source='D:\\course.xlsx';Extended Properties=Excel 8.0 ;");
                // OleDbDataAdapter ad = new OleDbDataAdapter();
                con.Open();
                String sql = "UPDATE[sheet1$] SET[Name] = '',[Year]= '',[Semester]= '',[Credit]= '' WHERE[Name] = '" + x + "' ";
                OleDbCommand cmd = new OleDbCommand(sql, con);
                cmd.ExecuteNonQuery();
                con.Close();

            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            MessageBox.Show("Record Deleted Successfully !");
        }
    }
}
