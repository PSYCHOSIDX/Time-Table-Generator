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
using System.Xml.Linq;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace TimeTableGeneratorDbce
{
    public partial class generateTable : Form
    {
        public generateTable()
        {
            InitializeComponent();
            Fillcombo();
            Fillcombo2();
            
        }
// dropdown for Courses | subjects
        void Fillcombo()
        {
            try
            {
                OleDbConnection con = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; Data Source='D:\\course.xlsx';Extended Properties=Excel 8.0 ;");
                con.Open();
                String sql = "select * from [sheet1$]";
                //String sql = "select * from [Sheet3$] where [day] = 'MON'";
                OleDbCommand cmd = new OleDbCommand(sql, con);
                OleDbDataReader Reader = cmd.ExecuteReader();
               while(Reader.Read())
                {
                    String name = Reader[0].ToString();
                    comboBox1.Items.Add( name);
                    comboBox2.Items.Add(name);
                    comboBox3.Items.Add(name);
                    comboBox4.Items.Add(name);
                    comboBox5.Items.Add(name);
                    comboBox6.Items.Add(name);

                    comboBox13.Items.Add(name);
                    comboBox14.Items.Add(name);
                    comboBox15.Items.Add(name);
                    comboBox16.Items.Add(name);
                    comboBox17.Items.Add(name);
                    comboBox18.Items.Add(name);

                    comboBox25.Items.Add(name);
                    comboBox26.Items.Add(name);
                    comboBox27.Items.Add(name);
                    comboBox28.Items.Add(name);
                    comboBox29.Items.Add(name);
                    comboBox30.Items.Add(name);

                    comboBox37.Items.Add(name);
                    comboBox38.Items.Add(name);
                    comboBox39.Items.Add(name);
                    comboBox40.Items.Add(name);
                    comboBox41.Items.Add(name);
                    comboBox42.Items.Add(name);

                    comboBox49.Items.Add(name);
                    comboBox50.Items.Add(name);
                    comboBox51.Items.Add(name);
                    comboBox52.Items.Add(name);
                    comboBox53.Items.Add(name);
                    comboBox54.Items.Add(name);

                    comboBox61.Items.Add(name);
                    comboBox62.Items.Add(name);
                    comboBox63.Items.Add(name);
                    comboBox64.Items.Add(name);
                    comboBox65.Items.Add(name);
                    comboBox66.Items.Add(name);
                }

                //cmd.ExecuteNonQuery();
                con.Close();

            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

           


        }

        //end of function Fillcombo()

        //dropdown for teachers | instructors
        void Fillcombo2()
        {
            try
            {
                OleDbConnection con = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; Data Source='D:\\course.xlsx';Extended Properties=Excel 8.0 ;");
                con.Open();
                String sql = "select * from [sheet2$]";
                OleDbCommand cmd = new OleDbCommand(sql, con);
                OleDbDataReader Reader = cmd.ExecuteReader();
                while (Reader.Read())
                {
                    String name = Reader[0].ToString();
                    comboBox7.Items.Add(name);
                    comboBox8.Items.Add(name);
                    comboBox9.Items.Add(name);
                    comboBox10.Items.Add(name);
                    comboBox11.Items.Add(name);
                    comboBox12.Items.Add(name);

                    comboBox19.Items.Add(name);
                    comboBox20.Items.Add(name);
                    comboBox21.Items.Add(name);
                    comboBox22.Items.Add(name);
                    comboBox23.Items.Add(name);
                    comboBox24.Items.Add(name);

                    comboBox31.Items.Add(name);
                    comboBox32.Items.Add(name);
                    comboBox33.Items.Add(name);
                    comboBox34.Items.Add(name);
                    comboBox35.Items.Add(name);
                    comboBox36.Items.Add(name);

                    comboBox43.Items.Add(name);
                    comboBox44.Items.Add(name);
                    comboBox45.Items.Add(name);
                    comboBox46.Items.Add(name);
                    comboBox47.Items.Add(name);
                    comboBox48.Items.Add(name);

                    comboBox55.Items.Add(name);
                    comboBox56.Items.Add(name);
                    comboBox57.Items.Add(name);
                    comboBox58.Items.Add(name);
                    comboBox59.Items.Add(name);
                    comboBox60.Items.Add(name);

                    comboBox67.Items.Add(name);
                    comboBox68.Items.Add(name);
                    comboBox69.Items.Add(name);
                    comboBox70.Items.Add(name);
                    comboBox71.Items.Add(name);
                    comboBox72.Items.Add(name);
                }

                //cmd.ExecuteNonQuery();
                con.Close();

            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }




        } //end of function Fillcombo2()


        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
           
        }

        private void button1_Click(object sender, EventArgs e)
{
            //start of input data for courses
            String a = comboBox1.Text;
            String b = comboBox2.Text;
            String c = comboBox3.Text;
            String d = comboBox4.Text;
            String f = comboBox5.Text;
            String g = comboBox6.Text;

            String h = comboBox13.Text;
            String i = comboBox14.Text;
            String j = comboBox15.Text;
            String k = comboBox16.Text;
            String l = comboBox17.Text;
            String m = comboBox18.Text;

            String n = comboBox25.Text;
            String o = comboBox26.Text;
            String p = comboBox27.Text;
            String q = comboBox28.Text;
            String r = comboBox29.Text;
            String s = comboBox30.Text;

            String t = comboBox37.Text;
            String u = comboBox38.Text;
            String v = comboBox39.Text;
            String w = comboBox40.Text;
            String x = comboBox41.Text;
            String y = comboBox42.Text;

            String z= comboBox49.Text;
            String z2 = comboBox50.Text;
            String z3 = comboBox51.Text;
            String z4 = comboBox52.Text;
            String z5 = comboBox53.Text;
            String z6 = comboBox54.Text;

            String z7 = comboBox61.Text;
            String z8 = comboBox62.Text;
            String z9 = comboBox63.Text;
            String z10 = comboBox64.Text;
            String z11 = comboBox65.Text;
            String z12 = comboBox66.Text;
            //end

            //start of input data for teachers

            String x1 = comboBox7.Text;
            String x2 = comboBox8.Text;
            String x3 = comboBox9.Text;
            String x4 = comboBox10.Text;
            String x5 = comboBox11.Text;
            String x6 = comboBox12.Text;

            String x7 = comboBox19.Text;
            String x8= comboBox20.Text;
            String x9 = comboBox21.Text;
            String x10 = comboBox22.Text;
            String x11 = comboBox23.Text;
            String x12 = comboBox24.Text;

            String x13 = comboBox31.Text;
            String x14 = comboBox32.Text;
            String x15 = comboBox33.Text;
            String x16 = comboBox34.Text;
            String x17 = comboBox35.Text;
            String x18 = comboBox36.Text;

            String x19 = comboBox43.Text;
            String x20 = comboBox44.Text;
            String x21 = comboBox45.Text;
            String x22 = comboBox46.Text;
            String x23 = comboBox47.Text;
            String x24 = comboBox48.Text;

            String x25 = comboBox55.Text;
            String x26 = comboBox56.Text;
            String x27 = comboBox57.Text;
            String x28 = comboBox58.Text;
            String x29 = comboBox59.Text;
            String x30 = comboBox60.Text;

            String x31 = comboBox67.Text;
            String x32 = comboBox68.Text;
            String x33 = comboBox69.Text;
            String x34 = comboBox70.Text;
            String x35 = comboBox71.Text;
            String x36 = comboBox72.Text;

            // end

            //mon
            try
            {
                OleDbConnection con = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; Data Source='D:\\course.xlsx';Extended Properties=Excel 8.0 ;");
                // OleDbDataAdapter ad = new OleDbDataAdapter();
                con.Open();
                
               String sql = "UPDATE [Sheet3$] SET[sub]='" + a + "',[sub1]='" + b + "',[sub2]='" + c + "',[sub3]='" + d + "',[sub4]='" + f + "',[sub5]='" + g + "' WHERE [day]='MON' ";
                OleDbCommand cmd = new OleDbCommand(sql, con);
                cmd.ExecuteNonQuery();
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            try
            {
                OleDbConnection con = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; Data Source='D:\\course.xlsx';Extended Properties=Excel 8.0 ;");
                // OleDbDataAdapter ad = new OleDbDataAdapter();
                con.Open();

                String sql = "UPDATE [Sheet3$] SET[sub]='" + x1 + "',[sub1]='" + x2 + "',[sub2]='" + x3 + "',[sub3]='" + x4 + "',[sub4]='" + x5+ "',[sub5]='" + x6 + "' WHERE [day]='1' ";
                OleDbCommand cmd = new OleDbCommand(sql, con);
                cmd.ExecuteNonQuery();
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            //tue
            try
            {
                OleDbConnection con = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; Data Source='D:\\course.xlsx';Extended Properties=Excel 8.0 ;");
                // OleDbDataAdapter ad = new OleDbDataAdapter();
                con.Open();

                String sql = "UPDATE [Sheet3$] SET[sub]='" + h + "',[sub1]='" + i + "',[sub2]='" + j + "',[sub3]='" + k + "',[sub4]='" + l + "',[sub5]='" + m + "' WHERE [day]='TUE' ";
                OleDbCommand cmd = new OleDbCommand(sql, con);
                cmd.ExecuteNonQuery();
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            try
            {
                OleDbConnection con = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; Data Source='D:\\course.xlsx';Extended Properties=Excel 8.0 ;");
                // OleDbDataAdapter ad = new OleDbDataAdapter();
                con.Open();

                String sql = "UPDATE [Sheet3$] SET[sub]='" + x7 + "',[sub1]='" + x8 + "',[sub2]='" + x9 + "',[sub3]='" + x10 + "',[sub4]='" + x11+ "',[sub5]='" + x12 + "' WHERE [day]='2' ";
                OleDbCommand cmd = new OleDbCommand(sql, con);
                cmd.ExecuteNonQuery();
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            //wed
            try
            {
                OleDbConnection con = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; Data Source='D:\\course.xlsx';Extended Properties=Excel 8.0 ;");
                // OleDbDataAdapter ad = new OleDbDataAdapter();
                con.Open();

                String sql = "UPDATE [Sheet3$] SET[sub]='" + n + "',[sub1]='" + o + "',[sub2]='" + p + "',[sub3]='" + q + "',[sub4]='" + r + "',[sub5]='" + s + "' WHERE [day]='WED' ";
                OleDbCommand cmd = new OleDbCommand(sql, con);
                cmd.ExecuteNonQuery();
                con.Close();
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            try
            {
                OleDbConnection con = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; Data Source='D:\\course.xlsx';Extended Properties=Excel 8.0 ;");
                // OleDbDataAdapter ad = new OleDbDataAdapter();
                con.Open();

                String sql = "UPDATE [Sheet3$] SET[sub]='" + x13 + "',[sub1]='" + x14 + "',[sub2]='" + x15 + "',[sub3]='" + x16 + "',[sub4]='" + x17 + "',[sub5]='" + x18 + "' WHERE [d]='3' ";
                OleDbCommand cmd = new OleDbCommand(sql, con);
                cmd.ExecuteNonQuery();
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }


            //thu
            try
            {
                OleDbConnection con = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; Data Source='D:\\course.xlsx';Extended Properties=Excel 8.0 ;");
                // OleDbDataAdapter ad = new OleDbDataAdapter();
                con.Open();

                String sql = "UPDATE [Sheet3$] SET[sub]='" + t + "',[sub1]='" + u + "',[sub2]='" + v + "',[sub3]='" + w + "',[sub4]='" + x + "',[sub5]='" + y + "' WHERE [day]='THU' ";
                OleDbCommand cmd = new OleDbCommand(sql, con);
                cmd.ExecuteNonQuery();
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            try
            {
                OleDbConnection con = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; Data Source='D:\\course.xlsx';Extended Properties=Excel 8.0 ;");
                // OleDbDataAdapter ad = new OleDbDataAdapter();
                con.Open();

                String sql = "UPDATE [Sheet3$] SET[sub]='" + x19 + "',[sub1]='" + x20 + "',[sub2]='" + x21 + "',[sub3]='" + x22 + "',[sub4]='" + x23 + "',[sub5]='" + x24 + "' WHERE [d]='4' ";
                OleDbCommand cmd = new OleDbCommand(sql, con);
                cmd.ExecuteNonQuery();
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }


            //fri
            try
            {
                OleDbConnection con = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; Data Source='D:\\course.xlsx';Extended Properties=Excel 8.0 ;");
                // OleDbDataAdapter ad = new OleDbDataAdapter();
                con.Open();

                String sql = "UPDATE [Sheet3$] SET[sub]='" + z + "',[sub1]='" + z2 + "',[sub2]='" + z3 + "',[sub3]='" + z4 + "',[sub4]='" + z5 + "',[sub5]='" + z6 + "' WHERE [day]='FRI' ";
                OleDbCommand cmd = new OleDbCommand(sql, con);
                cmd.ExecuteNonQuery();
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            try
            {
                OleDbConnection con = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; Data Source='D:\\course.xlsx';Extended Properties=Excel 8.0 ;");
                // OleDbDataAdapter ad = new OleDbDataAdapter();
                con.Open();

                String sql = "UPDATE [Sheet3$] SET[sub]='" + x25 + "',[sub1]='" + x26 + "',[sub2]='" + x27 + "',[sub3]='" + x28 + "',[sub4]='" + x29 + "',[sub5]='" + x30 + "' WHERE [d]='5' ";
                OleDbCommand cmd = new OleDbCommand(sql, con);
                cmd.ExecuteNonQuery();
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }


            //sat
            try
            {
                OleDbConnection con = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; Data Source='D:\\course.xlsx';Extended Properties=Excel 8.0 ;");
                // OleDbDataAdapter ad = new OleDbDataAdapter();
                con.Open();

                String sql = "UPDATE [Sheet3$] SET[sub]='" + z7 + "',[sub1]='" + z8 + "',[sub2]='" + z9 + "',[sub3]='" + z10 + "',[sub4]='" + z11 + "',[sub5]='" + z12+ "' WHERE [day]='SAT' ";
                OleDbCommand cmd = new OleDbCommand(sql, con);
                cmd.ExecuteNonQuery();
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            try
            {
                OleDbConnection con = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; Data Source='D:\\course.xlsx';Extended Properties=Excel 8.0 ;");
                // OleDbDataAdapter ad = new OleDbDataAdapter();
                con.Open();

                String sql = "UPDATE [Sheet3$] SET[sub]='" + x31 + "',[sub1]='" + x32 + "',[sub2]='" + x33 + "',[sub3]='" + x34 + "',[sub4]='" + x35 + "',[sub5]='" + x36 + "' WHERE [d]='6' ";
                OleDbCommand cmd = new OleDbCommand(sql, con);
                cmd.ExecuteNonQuery();
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }


            MessageBox.Show("Process Successful !");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //start of input data for courses
            String a = comboBox1.Text;
            String b = comboBox2.Text;
            String c = comboBox3.Text;
            String d = comboBox4.Text;
            String f = comboBox5.Text;
            String g = comboBox6.Text;

            String h = comboBox13.Text;
            String i = comboBox14.Text;
            String j = comboBox15.Text;
            String k = comboBox16.Text;
            String l = comboBox17.Text;
            String m = comboBox18.Text;

            String n = comboBox25.Text;
            String o = comboBox26.Text;
            String p = comboBox27.Text;
            String q = comboBox28.Text;
            String r = comboBox29.Text;
            String s = comboBox30.Text;

            String t = comboBox37.Text;
            String u = comboBox38.Text;
            String v = comboBox39.Text;
            String w = comboBox40.Text;
            String x = comboBox41.Text;
            String y = comboBox42.Text;

            String z = comboBox49.Text;
            String z2 = comboBox50.Text;
            String z3 = comboBox51.Text;
            String z4 = comboBox52.Text;
            String z5 = comboBox53.Text;
            String z6 = comboBox54.Text;

            String z7 = comboBox61.Text;
            String z8 = comboBox62.Text;
            String z9 = comboBox63.Text;
            String z10 = comboBox64.Text;
            String z11 = comboBox65.Text;
            String z12 = comboBox66.Text;
            //end

            //start of input data for teachers

            String x1 = comboBox7.Text;
            String x2 = comboBox8.Text;
            String x3 = comboBox9.Text;
            String x4 = comboBox10.Text;
            String x5 = comboBox11.Text;
            String x6 = comboBox12.Text;

            String x7 = comboBox19.Text;
            String x8 = comboBox20.Text;
            String x9 = comboBox21.Text;
            String x10 = comboBox22.Text;
            String x11 = comboBox23.Text;
            String x12 = comboBox24.Text;

            String x13 = comboBox31.Text;
            String x14 = comboBox32.Text;
            String x15 = comboBox33.Text;
            String x16 = comboBox34.Text;
            String x17 = comboBox35.Text;
            String x18 = comboBox36.Text;

            String x19 = comboBox43.Text;
            String x20 = comboBox44.Text;
            String x21 = comboBox45.Text;
            String x22 = comboBox46.Text;
            String x23 = comboBox47.Text;
            String x24 = comboBox48.Text;

            String x25 = comboBox55.Text;
            String x26 = comboBox56.Text;
            String x27 = comboBox57.Text;
            String x28 = comboBox58.Text;
            String x29 = comboBox59.Text;
            String x30 = comboBox60.Text;

            String x31 = comboBox67.Text;
            String x32 = comboBox68.Text;
            String x33 = comboBox69.Text;
            String x34 = comboBox70.Text;
            String x35 = comboBox71.Text;
            String x36 = comboBox72.Text;

            // end

            //mon
            try
            {
                OleDbConnection con = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; Data Source='D:\\course.xlsx';Extended Properties=Excel 8.0 ;");
                // OleDbDataAdapter ad = new OleDbDataAdapter();
                con.Open();

                String sql = "UPDATE [Sheet4$] SET[sub]='" + a + "',[sub1]='" + b + "',[sub2]='" + c + "',[sub3]='" + d + "',[sub4]='" + f + "',[sub5]='" + g + "' WHERE [day]='MON' ";
                OleDbCommand cmd = new OleDbCommand(sql, con);
                cmd.ExecuteNonQuery();
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            try
            {
                OleDbConnection con = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; Data Source='D:\\course.xlsx';Extended Properties=Excel 8.0 ;");
                // OleDbDataAdapter ad = new OleDbDataAdapter();
                con.Open();

                String sql = "UPDATE [Sheet4$] SET[sub]='" + x1 + "',[sub1]='" + x2 + "',[sub2]='" + x3 + "',[sub3]='" + x4 + "',[sub4]='" + x5 + "',[sub5]='" + x6 + "' WHERE [day]='1' ";
                OleDbCommand cmd = new OleDbCommand(sql, con);
                cmd.ExecuteNonQuery();
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            //tue
            try
            {
                OleDbConnection con = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; Data Source='D:\\course.xlsx';Extended Properties=Excel 8.0 ;");
                // OleDbDataAdapter ad = new OleDbDataAdapter();
                con.Open();

                String sql = "UPDATE [Sheet4$] SET[sub]='" + h + "',[sub1]='" + i + "',[sub2]='" + j + "',[sub3]='" + k + "',[sub4]='" + l + "',[sub5]='" + m + "' WHERE [day]='TUE' ";
                OleDbCommand cmd = new OleDbCommand(sql, con);
                cmd.ExecuteNonQuery();
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            try
            {
                OleDbConnection con = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; Data Source='D:\\course.xlsx';Extended Properties=Excel 8.0 ;");
                // OleDbDataAdapter ad = new OleDbDataAdapter();
                con.Open();

                String sql = "UPDATE [Sheet4$] SET[sub]='" + x7 + "',[sub1]='" + x8 + "',[sub2]='" + x9 + "',[sub3]='" + x10 + "',[sub4]='" + x11 + "',[sub5]='" + x12 + "' WHERE [day]='2' ";
                OleDbCommand cmd = new OleDbCommand(sql, con);
                cmd.ExecuteNonQuery();
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            //wed
            try
            {
                OleDbConnection con = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; Data Source='D:\\course.xlsx';Extended Properties=Excel 8.0 ;");
                // OleDbDataAdapter ad = new OleDbDataAdapter();
                con.Open();

                String sql = "UPDATE [Sheet4$] SET[sub]='" + n + "',[sub1]='" + o + "',[sub2]='" + p + "',[sub3]='" + q + "',[sub4]='" + r + "',[sub5]='" + s + "' WHERE [day]='WED' ";
                OleDbCommand cmd = new OleDbCommand(sql, con);
                cmd.ExecuteNonQuery();
                con.Close();
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            try
            {
                OleDbConnection con = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; Data Source='D:\\course.xlsx';Extended Properties=Excel 8.0 ;");
                // OleDbDataAdapter ad = new OleDbDataAdapter();
                con.Open();

                String sql = "UPDATE [Sheet4$] SET[sub]='" + x13 + "',[sub1]='" + x14 + "',[sub2]='" + x15 + "',[sub3]='" + x16 + "',[sub4]='" + x17 + "',[sub5]='" + x18 + "' WHERE [d]='3' ";
                OleDbCommand cmd = new OleDbCommand(sql, con);
                cmd.ExecuteNonQuery();
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }


            //thu
            try
            {
                OleDbConnection con = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; Data Source='D:\\course.xlsx';Extended Properties=Excel 8.0 ;");
                // OleDbDataAdapter ad = new OleDbDataAdapter();
                con.Open();

                String sql = "UPDATE [Sheet4$] SET[sub]='" + t + "',[sub1]='" + u + "',[sub2]='" + v + "',[sub3]='" + w + "',[sub4]='" + x + "',[sub5]='" + y + "' WHERE [day]='THU' ";
                OleDbCommand cmd = new OleDbCommand(sql, con);
                cmd.ExecuteNonQuery();
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            try
            {
                OleDbConnection con = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; Data Source='D:\\course.xlsx';Extended Properties=Excel 8.0 ;");
                // OleDbDataAdapter ad = new OleDbDataAdapter();
                con.Open();

                String sql = "UPDATE [Sheet4$] SET[sub]='" + x19 + "',[sub1]='" + x20 + "',[sub2]='" + x21 + "',[sub3]='" + x22 + "',[sub4]='" + x23 + "',[sub5]='" + x24 + "' WHERE [d]='4' ";
                OleDbCommand cmd = new OleDbCommand(sql, con);
                cmd.ExecuteNonQuery();
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }


            //fri
            try
            {
                OleDbConnection con = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; Data Source='D:\\course.xlsx';Extended Properties=Excel 8.0 ;");
                // OleDbDataAdapter ad = new OleDbDataAdapter();
                con.Open();

                String sql = "UPDATE [Sheet4$] SET[sub]='" + z + "',[sub1]='" + z2 + "',[sub2]='" + z3 + "',[sub3]='" + z4 + "',[sub4]='" + z5 + "',[sub5]='" + z6 + "' WHERE [day]='FRI' ";
                OleDbCommand cmd = new OleDbCommand(sql, con);
                cmd.ExecuteNonQuery();
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            try
            {
                OleDbConnection con = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; Data Source='D:\\course.xlsx';Extended Properties=Excel 8.0 ;");
                // OleDbDataAdapter ad = new OleDbDataAdapter();
                con.Open();

                String sql = "UPDATE [Sheet4$] SET[sub]='" + x25 + "',[sub1]='" + x26 + "',[sub2]='" + x27 + "',[sub3]='" + x28 + "',[sub4]='" + x29 + "',[sub5]='" + x30 + "' WHERE [d]='5' ";
                OleDbCommand cmd = new OleDbCommand(sql, con);
                cmd.ExecuteNonQuery();
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }


            //sat
            try
            {
                OleDbConnection con = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; Data Source='D:\\course.xlsx';Extended Properties=Excel 8.0 ;");
                // OleDbDataAdapter ad = new OleDbDataAdapter();
                con.Open();

                String sql = "UPDATE [Sheet4$] SET[sub]='" + z7 + "',[sub1]='" + z8 + "',[sub2]='" + z9 + "',[sub3]='" + z10 + "',[sub4]='" + z11 + "',[sub5]='" + z12 + "' WHERE [day]='SAT' ";
                OleDbCommand cmd = new OleDbCommand(sql, con);
                cmd.ExecuteNonQuery();
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            try
            {
                OleDbConnection con = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; Data Source='D:\\course.xlsx';Extended Properties=Excel 8.0 ;");
                // OleDbDataAdapter ad = new OleDbDataAdapter();
                con.Open();

                String sql = "UPDATE [Sheet4$] SET[sub]='" + x31 + "',[sub1]='" + x32 + "',[sub2]='" + x33 + "',[sub3]='" + x34 + "',[sub4]='" + x35 + "',[sub5]='" + x36 + "' WHERE [d]='6' ";
                OleDbCommand cmd = new OleDbCommand(sql, con);
                cmd.ExecuteNonQuery();
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }


            MessageBox.Show("Process Successful !");

        }

        private void button3_Click(object sender, EventArgs e)
        {
            //start of input data for courses
            String a = comboBox1.Text;
            String b = comboBox2.Text;
            String c = comboBox3.Text;
            String d = comboBox4.Text;
            String f = comboBox5.Text;
            String g = comboBox6.Text;

            String h = comboBox13.Text;
            String i = comboBox14.Text;
            String j = comboBox15.Text;
            String k = comboBox16.Text;
            String l = comboBox17.Text;
            String m = comboBox18.Text;

            String n = comboBox25.Text;
            String o = comboBox26.Text;
            String p = comboBox27.Text;
            String q = comboBox28.Text;
            String r = comboBox29.Text;
            String s = comboBox30.Text;

            String t = comboBox37.Text;
            String u = comboBox38.Text;
            String v = comboBox39.Text;
            String w = comboBox40.Text;
            String x = comboBox41.Text;
            String y = comboBox42.Text;

            String z = comboBox49.Text;
            String z2 = comboBox50.Text;
            String z3 = comboBox51.Text;
            String z4 = comboBox52.Text;
            String z5 = comboBox53.Text;
            String z6 = comboBox54.Text;

            String z7 = comboBox61.Text;
            String z8 = comboBox62.Text;
            String z9 = comboBox63.Text;
            String z10 = comboBox64.Text;
            String z11 = comboBox65.Text;
            String z12 = comboBox66.Text;
            //end

            //start of input data for teachers

            String x1 = comboBox7.Text;
            String x2 = comboBox8.Text;
            String x3 = comboBox9.Text;
            String x4 = comboBox10.Text;
            String x5 = comboBox11.Text;
            String x6 = comboBox12.Text;

            String x7 = comboBox19.Text;
            String x8 = comboBox20.Text;
            String x9 = comboBox21.Text;
            String x10 = comboBox22.Text;
            String x11 = comboBox23.Text;
            String x12 = comboBox24.Text;

            String x13 = comboBox31.Text;
            String x14 = comboBox32.Text;
            String x15 = comboBox33.Text;
            String x16 = comboBox34.Text;
            String x17 = comboBox35.Text;
            String x18 = comboBox36.Text;

            String x19 = comboBox43.Text;
            String x20 = comboBox44.Text;
            String x21 = comboBox45.Text;
            String x22 = comboBox46.Text;
            String x23 = comboBox47.Text;
            String x24 = comboBox48.Text;

            String x25 = comboBox55.Text;
            String x26 = comboBox56.Text;
            String x27 = comboBox57.Text;
            String x28 = comboBox58.Text;
            String x29 = comboBox59.Text;
            String x30 = comboBox60.Text;

            String x31 = comboBox67.Text;
            String x32 = comboBox68.Text;
            String x33 = comboBox69.Text;
            String x34 = comboBox70.Text;
            String x35 = comboBox71.Text;
            String x36 = comboBox72.Text;

            // end

            //mon
            try
            {
                OleDbConnection con = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; Data Source='D:\\course.xlsx';Extended Properties=Excel 8.0 ;");
                // OleDbDataAdapter ad = new OleDbDataAdapter();
                con.Open();

                String sql = "UPDATE [Sheet5$] SET[sub]='" + a + "',[sub1]='" + b + "',[sub2]='" + c + "',[sub3]='" + d + "',[sub4]='" + f + "',[sub5]='" + g + "' WHERE [day]='MON' ";
                OleDbCommand cmd = new OleDbCommand(sql, con);
                cmd.ExecuteNonQuery();
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            try
            {
                OleDbConnection con = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; Data Source='D:\\course.xlsx';Extended Properties=Excel 8.0 ;");
                // OleDbDataAdapter ad = new OleDbDataAdapter();
                con.Open();

                String sql = "UPDATE [Sheet5$] SET[sub]='" + x1 + "',[sub1]='" + x2 + "',[sub2]='" + x3 + "',[sub3]='" + x4 + "',[sub4]='" + x5 + "',[sub5]='" + x6 + "' WHERE [day]='1' ";
                OleDbCommand cmd = new OleDbCommand(sql, con);
                cmd.ExecuteNonQuery();
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            //tue
            try
            {
                OleDbConnection con = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; Data Source='D:\\course.xlsx';Extended Properties=Excel 8.0 ;");
                // OleDbDataAdapter ad = new OleDbDataAdapter();
                con.Open();

                String sql = "UPDATE [Sheet5$] SET[sub]='" + h + "',[sub1]='" + i + "',[sub2]='" + j + "',[sub3]='" + k + "',[sub4]='" + l + "',[sub5]='" + m + "' WHERE [day]='TUE' ";
                OleDbCommand cmd = new OleDbCommand(sql, con);
                cmd.ExecuteNonQuery();
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            try
            {
                OleDbConnection con = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; Data Source='D:\\course.xlsx';Extended Properties=Excel 8.0 ;");
                // OleDbDataAdapter ad = new OleDbDataAdapter();
                con.Open();

                String sql = "UPDATE [Sheet5$] SET[sub]='" + x7 + "',[sub1]='" + x8 + "',[sub2]='" + x9 + "',[sub3]='" + x10 + "',[sub4]='" + x11 + "',[sub5]='" + x12 + "' WHERE [day]='2' ";
                OleDbCommand cmd = new OleDbCommand(sql, con);
                cmd.ExecuteNonQuery();
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            //wed
            try
            {
                OleDbConnection con = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; Data Source='D:\\course.xlsx';Extended Properties=Excel 8.0 ;");
                // OleDbDataAdapter ad = new OleDbDataAdapter();
                con.Open();

                String sql = "UPDATE [Sheet5$] SET[sub]='" + n + "',[sub1]='" + o + "',[sub2]='" + p + "',[sub3]='" + q + "',[sub4]='" + r + "',[sub5]='" + s + "' WHERE [day]='WED' ";
                OleDbCommand cmd = new OleDbCommand(sql, con);
                cmd.ExecuteNonQuery();
                con.Close();
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            try
            {
                OleDbConnection con = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; Data Source='D:\\course.xlsx';Extended Properties=Excel 8.0 ;");
                // OleDbDataAdapter ad = new OleDbDataAdapter();
                con.Open();

                String sql = "UPDATE [Sheet5$] SET[sub]='" + x13 + "',[sub1]='" + x14 + "',[sub2]='" + x15 + "',[sub3]='" + x16 + "',[sub4]='" + x17 + "',[sub5]='" + x18 + "' WHERE [d]='3' ";
                OleDbCommand cmd = new OleDbCommand(sql, con);
                cmd.ExecuteNonQuery();
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }


            //thu
            try
            {
                OleDbConnection con = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; Data Source='D:\\course.xlsx';Extended Properties=Excel 8.0 ;");
                // OleDbDataAdapter ad = new OleDbDataAdapter();
                con.Open();

                String sql = "UPDATE [Sheet5$] SET[sub]='" + t + "',[sub1]='" + u + "',[sub2]='" + v + "',[sub3]='" + w + "',[sub4]='" + x + "',[sub5]='" + y + "' WHERE [day]='THU' ";
                OleDbCommand cmd = new OleDbCommand(sql, con);
                cmd.ExecuteNonQuery();
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            try
            {
                OleDbConnection con = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; Data Source='D:\\course.xlsx';Extended Properties=Excel 8.0 ;");
                // OleDbDataAdapter ad = new OleDbDataAdapter();
                con.Open();

                String sql = "UPDATE [Sheet5$] SET[sub]='" + x19 + "',[sub1]='" + x20 + "',[sub2]='" + x21 + "',[sub3]='" + x22 + "',[sub4]='" + x23 + "',[sub5]='" + x24 + "' WHERE [d]='4' ";
                OleDbCommand cmd = new OleDbCommand(sql, con);
                cmd.ExecuteNonQuery();
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }


            //fri
            try
            {
                OleDbConnection con = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; Data Source='D:\\course.xlsx';Extended Properties=Excel 8.0 ;");
                // OleDbDataAdapter ad = new OleDbDataAdapter();
                con.Open();

                String sql = "UPDATE [Sheet5$] SET[sub]='" + z + "',[sub1]='" + z2 + "',[sub2]='" + z3 + "',[sub3]='" + z4 + "',[sub4]='" + z5 + "',[sub5]='" + z6 + "' WHERE [day]='FRI' ";
                OleDbCommand cmd = new OleDbCommand(sql, con);
                cmd.ExecuteNonQuery();
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            try
            {
                OleDbConnection con = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; Data Source='D:\\course.xlsx';Extended Properties=Excel 8.0 ;");
                // OleDbDataAdapter ad = new OleDbDataAdapter();
                con.Open();

                String sql = "UPDATE [Sheet5$] SET[sub]='" + x25 + "',[sub1]='" + x26 + "',[sub2]='" + x27 + "',[sub3]='" + x28 + "',[sub4]='" + x29 + "',[sub5]='" + x30 + "' WHERE [d]='5' ";
                OleDbCommand cmd = new OleDbCommand(sql, con);
                cmd.ExecuteNonQuery();
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }


            //sat
            try
            {
                OleDbConnection con = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; Data Source='D:\\course.xlsx';Extended Properties=Excel 8.0 ;");
                // OleDbDataAdapter ad = new OleDbDataAdapter();
                con.Open();

                String sql = "UPDATE [Sheet5$] SET[sub]='" + z7 + "',[sub1]='" + z8 + "',[sub2]='" + z9 + "',[sub3]='" + z10 + "',[sub4]='" + z11 + "',[sub5]='" + z12 + "' WHERE [day]='SAT' ";
                OleDbCommand cmd = new OleDbCommand(sql, con);
                cmd.ExecuteNonQuery();
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            try
            {
                OleDbConnection con = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; Data Source='D:\\course.xlsx';Extended Properties=Excel 8.0 ;");
                // OleDbDataAdapter ad = new OleDbDataAdapter();
                con.Open();

                String sql = "UPDATE [Sheet5$] SET[sub]='" + x31 + "',[sub1]='" + x32 + "',[sub2]='" + x33 + "',[sub3]='" + x34 + "',[sub4]='" + x35 + "',[sub5]='" + x36 + "' WHERE [d]='6' ";
                OleDbCommand cmd = new OleDbCommand(sql, con);
                cmd.ExecuteNonQuery();
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }


            MessageBox.Show("Process Successful !");

        }

        private void button4_Click(object sender, EventArgs e)
        {
            //start of input data for courses
            String a = comboBox1.Text;
            String b = comboBox2.Text;
            String c = comboBox3.Text;
            String d = comboBox4.Text;
            String f = comboBox5.Text;
            String g = comboBox6.Text;

            String h = comboBox13.Text;
            String i = comboBox14.Text;
            String j = comboBox15.Text;
            String k = comboBox16.Text;
            String l = comboBox17.Text;
            String m = comboBox18.Text;

            String n = comboBox25.Text;
            String o = comboBox26.Text;
            String p = comboBox27.Text;
            String q = comboBox28.Text;
            String r = comboBox29.Text;
            String s = comboBox30.Text;

            String t = comboBox37.Text;
            String u = comboBox38.Text;
            String v = comboBox39.Text;
            String w = comboBox40.Text;
            String x = comboBox41.Text;
            String y = comboBox42.Text;

            String z = comboBox49.Text;
            String z2 = comboBox50.Text;
            String z3 = comboBox51.Text;
            String z4 = comboBox52.Text;
            String z5 = comboBox53.Text;
            String z6 = comboBox54.Text;

            String z7 = comboBox61.Text;
            String z8 = comboBox62.Text;
            String z9 = comboBox63.Text;
            String z10 = comboBox64.Text;
            String z11 = comboBox65.Text;
            String z12 = comboBox66.Text;
            //end

            //start of input data for teachers

            String x1 = comboBox7.Text;
            String x2 = comboBox8.Text;
            String x3 = comboBox9.Text;
            String x4 = comboBox10.Text;
            String x5 = comboBox11.Text;
            String x6 = comboBox12.Text;

            String x7 = comboBox19.Text;
            String x8 = comboBox20.Text;
            String x9 = comboBox21.Text;
            String x10 = comboBox22.Text;
            String x11 = comboBox23.Text;
            String x12 = comboBox24.Text;

            String x13 = comboBox31.Text;
            String x14 = comboBox32.Text;
            String x15 = comboBox33.Text;
            String x16 = comboBox34.Text;
            String x17 = comboBox35.Text;
            String x18 = comboBox36.Text;

            String x19 = comboBox43.Text;
            String x20 = comboBox44.Text;
            String x21 = comboBox45.Text;
            String x22 = comboBox46.Text;
            String x23 = comboBox47.Text;
            String x24 = comboBox48.Text;

            String x25 = comboBox55.Text;
            String x26 = comboBox56.Text;
            String x27 = comboBox57.Text;
            String x28 = comboBox58.Text;
            String x29 = comboBox59.Text;
            String x30 = comboBox60.Text;

            String x31 = comboBox67.Text;
            String x32 = comboBox68.Text;
            String x33 = comboBox69.Text;
            String x34 = comboBox70.Text;
            String x35 = comboBox71.Text;
            String x36 = comboBox72.Text;

            // end

            //mon
            try
            {
                OleDbConnection con = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; Data Source='D:\\course.xlsx';Extended Properties=Excel 8.0 ;");
                // OleDbDataAdapter ad = new OleDbDataAdapter();
                con.Open();

                String sql = "UPDATE [Sheet6$] SET[sub]='" + a + "',[sub1]='" + b + "',[sub2]='" + c + "',[sub3]='" + d + "',[sub4]='" + f + "',[sub5]='" + g + "' WHERE [day]='MON' ";
                OleDbCommand cmd = new OleDbCommand(sql, con);
                cmd.ExecuteNonQuery();
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            try
            {
                OleDbConnection con = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; Data Source='D:\\course.xlsx';Extended Properties=Excel 8.0 ;");
                // OleDbDataAdapter ad = new OleDbDataAdapter();
                con.Open();

                String sql = "UPDATE [Sheet6$] SET[sub]='" + x1 + "',[sub1]='" + x2 + "',[sub2]='" + x3 + "',[sub3]='" + x4 + "',[sub4]='" + x5 + "',[sub5]='" + x6 + "' WHERE [day]='1' ";
                OleDbCommand cmd = new OleDbCommand(sql, con);
                cmd.ExecuteNonQuery();
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            //tue
            try
            {
                OleDbConnection con = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; Data Source='D:\\course.xlsx';Extended Properties=Excel 8.0 ;");
                // OleDbDataAdapter ad = new OleDbDataAdapter();
                con.Open();

                String sql = "UPDATE [Sheet6$] SET[sub]='" + h + "',[sub1]='" + i + "',[sub2]='" + j + "',[sub3]='" + k + "',[sub4]='" + l + "',[sub5]='" + m + "' WHERE [day]='TUE' ";
                OleDbCommand cmd = new OleDbCommand(sql, con);
                cmd.ExecuteNonQuery();
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            try
            {
                OleDbConnection con = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; Data Source='D:\\course.xlsx';Extended Properties=Excel 8.0 ;");
                // OleDbDataAdapter ad = new OleDbDataAdapter();
                con.Open();

                String sql = "UPDATE [Sheet6$] SET[sub]='" + x7 + "',[sub1]='" + x8 + "',[sub2]='" + x9 + "',[sub3]='" + x10 + "',[sub4]='" + x11 + "',[sub5]='" + x12 + "' WHERE [day]='2' ";
                OleDbCommand cmd = new OleDbCommand(sql, con);
                cmd.ExecuteNonQuery();
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            //wed
            try
            {
                OleDbConnection con = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; Data Source='D:\\course.xlsx';Extended Properties=Excel 8.0 ;");
                // OleDbDataAdapter ad = new OleDbDataAdapter();
                con.Open();

                String sql = "UPDATE [Sheet6$] SET[sub]='" + n + "',[sub1]='" + o + "',[sub2]='" + p + "',[sub3]='" + q + "',[sub4]='" + r + "',[sub5]='" + s + "' WHERE [day]='WED' ";
                OleDbCommand cmd = new OleDbCommand(sql, con);
                cmd.ExecuteNonQuery();
                con.Close();
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            try
            {
                OleDbConnection con = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; Data Source='D:\\course.xlsx';Extended Properties=Excel 8.0 ;");
                // OleDbDataAdapter ad = new OleDbDataAdapter();
                con.Open();

                String sql = "UPDATE [Sheet6$] SET[sub]='" + x13 + "',[sub1]='" + x14 + "',[sub2]='" + x15 + "',[sub3]='" + x16 + "',[sub4]='" + x17 + "',[sub5]='" + x18 + "' WHERE [d]='3' ";
                OleDbCommand cmd = new OleDbCommand(sql, con);
                cmd.ExecuteNonQuery();
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }


            //thu
            try
            {
                OleDbConnection con = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; Data Source='D:\\course.xlsx';Extended Properties=Excel 8.0 ;");
                // OleDbDataAdapter ad = new OleDbDataAdapter();
                con.Open();

                String sql = "UPDATE [Sheet6$] SET[sub]='" + t + "',[sub1]='" + u + "',[sub2]='" + v + "',[sub3]='" + w + "',[sub4]='" + x + "',[sub5]='" + y + "' WHERE [day]='THU' ";
                OleDbCommand cmd = new OleDbCommand(sql, con);
                cmd.ExecuteNonQuery();
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            try
            {
                OleDbConnection con = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; Data Source='D:\\course.xlsx';Extended Properties=Excel 8.0 ;");
                // OleDbDataAdapter ad = new OleDbDataAdapter();
                con.Open();

                String sql = "UPDATE [Sheet6$] SET[sub]='" + x19 + "',[sub1]='" + x20 + "',[sub2]='" + x21 + "',[sub3]='" + x22 + "',[sub4]='" + x23 + "',[sub5]='" + x24 + "' WHERE [d]='4' ";
                OleDbCommand cmd = new OleDbCommand(sql, con);
                cmd.ExecuteNonQuery();
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }


            //fri
            try
            {
                OleDbConnection con = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; Data Source='D:\\course.xlsx';Extended Properties=Excel 8.0 ;");
                // OleDbDataAdapter ad = new OleDbDataAdapter();
                con.Open();

                String sql = "UPDATE [Sheet6$] SET[sub]='" + z + "',[sub1]='" + z2 + "',[sub2]='" + z3 + "',[sub3]='" + z4 + "',[sub4]='" + z5 + "',[sub5]='" + z6 + "' WHERE [day]='FRI' ";
                OleDbCommand cmd = new OleDbCommand(sql, con);
                cmd.ExecuteNonQuery();
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            try
            {
                OleDbConnection con = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; Data Source='D:\\course.xlsx';Extended Properties=Excel 8.0 ;");
                // OleDbDataAdapter ad = new OleDbDataAdapter();
                con.Open();

                String sql = "UPDATE [Sheet6$] SET[sub]='" + x25 + "',[sub1]='" + x26 + "',[sub2]='" + x27 + "',[sub3]='" + x28 + "',[sub4]='" + x29 + "',[sub5]='" + x30 + "' WHERE [d]='5' ";
                OleDbCommand cmd = new OleDbCommand(sql, con);
                cmd.ExecuteNonQuery();
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }


            //sat
            try
            {
                OleDbConnection con = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; Data Source='D:\\course.xlsx';Extended Properties=Excel 8.0 ;");
                // OleDbDataAdapter ad = new OleDbDataAdapter();
                con.Open();

                String sql = "UPDATE [Sheet6$] SET[sub]='" + z7 + "',[sub1]='" + z8 + "',[sub2]='" + z9 + "',[sub3]='" + z10 + "',[sub4]='" + z11 + "',[sub5]='" + z12 + "' WHERE [day]='SAT' ";
                OleDbCommand cmd = new OleDbCommand(sql, con);
                cmd.ExecuteNonQuery();
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            try
            {
                OleDbConnection con = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; Data Source='D:\\course.xlsx';Extended Properties=Excel 8.0 ;");
                // OleDbDataAdapter ad = new OleDbDataAdapter();
                con.Open();

                String sql = "UPDATE [Sheet6$] SET[sub]='" + x31 + "',[sub1]='" + x32 + "',[sub2]='" + x33 + "',[sub3]='" + x34 + "',[sub4]='" + x35 + "',[sub5]='" + x36 + "' WHERE [d]='6' ";
                OleDbCommand cmd = new OleDbCommand(sql, con);
                cmd.ExecuteNonQuery();
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }


            MessageBox.Show("Process Successful !");

        }

        private void button5_Click(object sender, EventArgs e)
        {
            //start of input data for courses
            String a = comboBox1.Text;
            String b = comboBox2.Text;
            String c = comboBox3.Text;
            String d = comboBox4.Text;
            String f = comboBox5.Text;
            String g = comboBox6.Text;

            String h = comboBox13.Text;
            String i = comboBox14.Text;
            String j = comboBox15.Text;
            String k = comboBox16.Text;
            String l = comboBox17.Text;
            String m = comboBox18.Text;

            String n = comboBox25.Text;
            String o = comboBox26.Text;
            String p = comboBox27.Text;
            String q = comboBox28.Text;
            String r = comboBox29.Text;
            String s = comboBox30.Text;

            String t = comboBox37.Text;
            String u = comboBox38.Text;
            String v = comboBox39.Text;
            String w = comboBox40.Text;
            String x = comboBox41.Text;
            String y = comboBox42.Text;

            String z = comboBox49.Text;
            String z2 = comboBox50.Text;
            String z3 = comboBox51.Text;
            String z4 = comboBox52.Text;
            String z5 = comboBox53.Text;
            String z6 = comboBox54.Text;

            String z7 = comboBox61.Text;
            String z8 = comboBox62.Text;
            String z9 = comboBox63.Text;
            String z10 = comboBox64.Text;
            String z11 = comboBox65.Text;
            String z12 = comboBox66.Text;
            //end

            //start of input data for teachers

            String x1 = comboBox7.Text;
            String x2 = comboBox8.Text;
            String x3 = comboBox9.Text;
            String x4 = comboBox10.Text;
            String x5 = comboBox11.Text;
            String x6 = comboBox12.Text;

            String x7 = comboBox19.Text;
            String x8 = comboBox20.Text;
            String x9 = comboBox21.Text;
            String x10 = comboBox22.Text;
            String x11 = comboBox23.Text;
            String x12 = comboBox24.Text;

            String x13 = comboBox31.Text;
            String x14 = comboBox32.Text;
            String x15 = comboBox33.Text;
            String x16 = comboBox34.Text;
            String x17 = comboBox35.Text;
            String x18 = comboBox36.Text;

            String x19 = comboBox43.Text;
            String x20 = comboBox44.Text;
            String x21 = comboBox45.Text;
            String x22 = comboBox46.Text;
            String x23 = comboBox47.Text;
            String x24 = comboBox48.Text;

            String x25 = comboBox55.Text;
            String x26 = comboBox56.Text;
            String x27 = comboBox57.Text;
            String x28 = comboBox58.Text;
            String x29 = comboBox59.Text;
            String x30 = comboBox60.Text;

            String x31 = comboBox67.Text;
            String x32 = comboBox68.Text;
            String x33 = comboBox69.Text;
            String x34 = comboBox70.Text;
            String x35 = comboBox71.Text;
            String x36 = comboBox72.Text;

            // end

            //mon
            try
            {
                OleDbConnection con = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; Data Source='D:\\course.xlsx';Extended Properties=Excel 8.0 ;");
                // OleDbDataAdapter ad = new OleDbDataAdapter();
                con.Open();

                String sql = "UPDATE [Sheet7$] SET[sub]='" + a + "',[sub1]='" + b + "',[sub2]='" + c + "',[sub3]='" + d + "',[sub4]='" + f + "',[sub5]='" + g + "' WHERE [day]='MON' ";
                OleDbCommand cmd = new OleDbCommand(sql, con);
                cmd.ExecuteNonQuery();
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            try
            {
                OleDbConnection con = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; Data Source='D:\\course.xlsx';Extended Properties=Excel 8.0 ;");
                // OleDbDataAdapter ad = new OleDbDataAdapter();
                con.Open();

                String sql = "UPDATE [Sheet7$] SET[sub]='" + x1 + "',[sub1]='" + x2 + "',[sub2]='" + x3 + "',[sub3]='" + x4 + "',[sub4]='" + x5 + "',[sub5]='" + x6 + "' WHERE [day]='1' ";
                OleDbCommand cmd = new OleDbCommand(sql, con);
                cmd.ExecuteNonQuery();
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            //tue
            try
            {
                OleDbConnection con = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; Data Source='D:\\course.xlsx';Extended Properties=Excel 8.0 ;");
                // OleDbDataAdapter ad = new OleDbDataAdapter();
                con.Open();

                String sql = "UPDATE [Sheet7$] SET[sub]='" + h + "',[sub1]='" + i + "',[sub2]='" + j + "',[sub3]='" + k + "',[sub4]='" + l + "',[sub5]='" + m + "' WHERE [day]='TUE' ";
                OleDbCommand cmd = new OleDbCommand(sql, con);
                cmd.ExecuteNonQuery();
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            try
            {
                OleDbConnection con = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; Data Source='D:\\course.xlsx';Extended Properties=Excel 8.0 ;");
                // OleDbDataAdapter ad = new OleDbDataAdapter();
                con.Open();

                String sql = "UPDATE [Sheet7$] SET[sub]='" + x7 + "',[sub1]='" + x8 + "',[sub2]='" + x9 + "',[sub3]='" + x10 + "',[sub4]='" + x11 + "',[sub5]='" + x12 + "' WHERE [day]='2' ";
                OleDbCommand cmd = new OleDbCommand(sql, con);
                cmd.ExecuteNonQuery();
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            //wed
            try
            {
                OleDbConnection con = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; Data Source='D:\\course.xlsx';Extended Properties=Excel 8.0 ;");
                // OleDbDataAdapter ad = new OleDbDataAdapter();
                con.Open();

                String sql = "UPDATE [Sheet7$] SET[sub]='" + n + "',[sub1]='" + o + "',[sub2]='" + p + "',[sub3]='" + q + "',[sub4]='" + r + "',[sub5]='" + s + "' WHERE [day]='WED' ";
                OleDbCommand cmd = new OleDbCommand(sql, con);
                cmd.ExecuteNonQuery();
                con.Close();
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            try
            {
                OleDbConnection con = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; Data Source='D:\\course.xlsx';Extended Properties=Excel 8.0 ;");
                // OleDbDataAdapter ad = new OleDbDataAdapter();
                con.Open();

                String sql = "UPDATE [Sheet7$] SET[sub]='" + x13 + "',[sub1]='" + x14 + "',[sub2]='" + x15 + "',[sub3]='" + x16 + "',[sub4]='" + x17 + "',[sub5]='" + x18 + "' WHERE [d]='3' ";
                OleDbCommand cmd = new OleDbCommand(sql, con);
                cmd.ExecuteNonQuery();
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }


            //thu
            try
            {
                OleDbConnection con = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; Data Source='D:\\course.xlsx';Extended Properties=Excel 8.0 ;");
                // OleDbDataAdapter ad = new OleDbDataAdapter();
                con.Open();

                String sql = "UPDATE [Sheet7$] SET[sub]='" + t + "',[sub1]='" + u + "',[sub2]='" + v + "',[sub3]='" + w + "',[sub4]='" + x + "',[sub5]='" + y + "' WHERE [day]='THU' ";
                OleDbCommand cmd = new OleDbCommand(sql, con);
                cmd.ExecuteNonQuery();
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            try
            {
                OleDbConnection con = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; Data Source='D:\\course.xlsx';Extended Properties=Excel 8.0 ;");
                // OleDbDataAdapter ad = new OleDbDataAdapter();
                con.Open();

                String sql = "UPDATE [Sheet7$] SET[sub]='" + x19 + "',[sub1]='" + x20 + "',[sub2]='" + x21 + "',[sub3]='" + x22 + "',[sub4]='" + x23 + "',[sub5]='" + x24 + "' WHERE [d]='4' ";
                OleDbCommand cmd = new OleDbCommand(sql, con);
                cmd.ExecuteNonQuery();
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }


            //fri
            try
            {
                OleDbConnection con = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; Data Source='D:\\course.xlsx';Extended Properties=Excel 8.0 ;");
                // OleDbDataAdapter ad = new OleDbDataAdapter();
                con.Open();

                String sql = "UPDATE [Sheet7$] SET[sub]='" + z + "',[sub1]='" + z2 + "',[sub2]='" + z3 + "',[sub3]='" + z4 + "',[sub4]='" + z5 + "',[sub5]='" + z6 + "' WHERE [day]='FRI' ";
                OleDbCommand cmd = new OleDbCommand(sql, con);
                cmd.ExecuteNonQuery();
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            try
            {
                OleDbConnection con = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; Data Source='D:\\course.xlsx';Extended Properties=Excel 8.0 ;");
                // OleDbDataAdapter ad = new OleDbDataAdapter();
                con.Open();

                String sql = "UPDATE [Sheet7$] SET[sub]='" + x25 + "',[sub1]='" + x26 + "',[sub2]='" + x27 + "',[sub3]='" + x28 + "',[sub4]='" + x29 + "',[sub5]='" + x30 + "' WHERE [d]='5' ";
                OleDbCommand cmd = new OleDbCommand(sql, con);
                cmd.ExecuteNonQuery();
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }


            //sat
            try
            {
                OleDbConnection con = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; Data Source='D:\\course.xlsx';Extended Properties=Excel 8.0 ;");
                // OleDbDataAdapter ad = new OleDbDataAdapter();
                con.Open();

                String sql = "UPDATE [Sheet7$] SET[sub]='" + z7 + "',[sub1]='" + z8 + "',[sub2]='" + z9 + "',[sub3]='" + z10 + "',[sub4]='" + z11 + "',[sub5]='" + z12 + "' WHERE [day]='SAT' ";
                OleDbCommand cmd = new OleDbCommand(sql, con);
                cmd.ExecuteNonQuery();
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            try
            {
                OleDbConnection con = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; Data Source='D:\\course.xlsx';Extended Properties=Excel 8.0 ;");
                // OleDbDataAdapter ad = new OleDbDataAdapter();
                con.Open();

                String sql = "UPDATE [Sheet7$] SET[sub]='" + x31 + "',[sub1]='" + x32 + "',[sub2]='" + x33 + "',[sub3]='" + x34 + "',[sub4]='" + x35 + "',[sub5]='" + x36 + "' WHERE [d]='6' ";
                OleDbCommand cmd = new OleDbCommand(sql, con);
                cmd.ExecuteNonQuery();
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }


            MessageBox.Show("Process Successful !");

        }

        private void button6_Click(object sender, EventArgs e)
        {
            //start of input data for courses
            String a = comboBox1.Text;
            String b = comboBox2.Text;
            String c = comboBox3.Text;
            String d = comboBox4.Text;
            String f = comboBox5.Text;
            String g = comboBox6.Text;

            String h = comboBox13.Text;
            String i = comboBox14.Text;
            String j = comboBox15.Text;
            String k = comboBox16.Text;
            String l = comboBox17.Text;
            String m = comboBox18.Text;

            String n = comboBox25.Text;
            String o = comboBox26.Text;
            String p = comboBox27.Text;
            String q = comboBox28.Text;
            String r = comboBox29.Text;
            String s = comboBox30.Text;

            String t = comboBox37.Text;
            String u = comboBox38.Text;
            String v = comboBox39.Text;
            String w = comboBox40.Text;
            String x = comboBox41.Text;
            String y = comboBox42.Text;

            String z = comboBox49.Text;
            String z2 = comboBox50.Text;
            String z3 = comboBox51.Text;
            String z4 = comboBox52.Text;
            String z5 = comboBox53.Text;
            String z6 = comboBox54.Text;

            String z7 = comboBox61.Text;
            String z8 = comboBox62.Text;
            String z9 = comboBox63.Text;
            String z10 = comboBox64.Text;
            String z11 = comboBox65.Text;
            String z12 = comboBox66.Text;
            //end

            //start of input data for teachers

            String x1 = comboBox7.Text;
            String x2 = comboBox8.Text;
            String x3 = comboBox9.Text;
            String x4 = comboBox10.Text;
            String x5 = comboBox11.Text;
            String x6 = comboBox12.Text;

            String x7 = comboBox19.Text;
            String x8 = comboBox20.Text;
            String x9 = comboBox21.Text;
            String x10 = comboBox22.Text;
            String x11 = comboBox23.Text;
            String x12 = comboBox24.Text;

            String x13 = comboBox31.Text;
            String x14 = comboBox32.Text;
            String x15 = comboBox33.Text;
            String x16 = comboBox34.Text;
            String x17 = comboBox35.Text;
            String x18 = comboBox36.Text;

            String x19 = comboBox43.Text;
            String x20 = comboBox44.Text;
            String x21 = comboBox45.Text;
            String x22 = comboBox46.Text;
            String x23 = comboBox47.Text;
            String x24 = comboBox48.Text;

            String x25 = comboBox55.Text;
            String x26 = comboBox56.Text;
            String x27 = comboBox57.Text;
            String x28 = comboBox58.Text;
            String x29 = comboBox59.Text;
            String x30 = comboBox60.Text;

            String x31 = comboBox67.Text;
            String x32 = comboBox68.Text;
            String x33 = comboBox69.Text;
            String x34 = comboBox70.Text;
            String x35 = comboBox71.Text;
            String x36 = comboBox72.Text;

            // end

            //mon
            try
            {
                OleDbConnection con = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; Data Source='D:\\course.xlsx';Extended Properties=Excel 8.0 ;");
                // OleDbDataAdapter ad = new OleDbDataAdapter();
                con.Open();

                String sql = "UPDATE [Sheet8$] SET[sub]='" + a + "',[sub1]='" + b + "',[sub2]='" + c + "',[sub3]='" + d + "',[sub4]='" + f + "',[sub5]='" + g + "' WHERE [day]='MON' ";
                OleDbCommand cmd = new OleDbCommand(sql, con);
                cmd.ExecuteNonQuery();
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            try
            {
                OleDbConnection con = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; Data Source='D:\\course.xlsx';Extended Properties=Excel 8.0 ;");
                // OleDbDataAdapter ad = new OleDbDataAdapter();
                con.Open();

                String sql = "UPDATE [Sheet8$] SET[sub]='" + x1 + "',[sub1]='" + x2 + "',[sub2]='" + x3 + "',[sub3]='" + x4 + "',[sub4]='" + x5 + "',[sub5]='" + x6 + "' WHERE [day]='1' ";
                OleDbCommand cmd = new OleDbCommand(sql, con);
                cmd.ExecuteNonQuery();
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            //tue
            try
            {
                OleDbConnection con = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; Data Source='D:\\course.xlsx';Extended Properties=Excel 8.0 ;");
                // OleDbDataAdapter ad = new OleDbDataAdapter();
                con.Open();

                String sql = "UPDATE [Sheet8$] SET[sub]='" + h + "',[sub1]='" + i + "',[sub2]='" + j + "',[sub3]='" + k + "',[sub4]='" + l + "',[sub5]='" + m + "' WHERE [day]='TUE' ";
                OleDbCommand cmd = new OleDbCommand(sql, con);
                cmd.ExecuteNonQuery();
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            try
            {
                OleDbConnection con = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; Data Source='D:\\course.xlsx';Extended Properties=Excel 8.0 ;");
                // OleDbDataAdapter ad = new OleDbDataAdapter();
                con.Open();

                String sql = "UPDATE [Sheet8$] SET[sub]='" + x7 + "',[sub1]='" + x8 + "',[sub2]='" + x9 + "',[sub3]='" + x10 + "',[sub4]='" + x11 + "',[sub5]='" + x12 + "' WHERE [day]='2' ";
                OleDbCommand cmd = new OleDbCommand(sql, con);
                cmd.ExecuteNonQuery();
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            //wed
            try
            {
                OleDbConnection con = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; Data Source='D:\\course.xlsx';Extended Properties=Excel 8.0 ;");
                // OleDbDataAdapter ad = new OleDbDataAdapter();
                con.Open();

                String sql = "UPDATE [Sheet8$] SET[sub]='" + n + "',[sub1]='" + o + "',[sub2]='" + p + "',[sub3]='" + q + "',[sub4]='" + r + "',[sub5]='" + s + "' WHERE [day]='WED' ";
                OleDbCommand cmd = new OleDbCommand(sql, con);
                cmd.ExecuteNonQuery();
                con.Close();
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            try
            {
                OleDbConnection con = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; Data Source='D:\\course.xlsx';Extended Properties=Excel 8.0 ;");
                // OleDbDataAdapter ad = new OleDbDataAdapter();
                con.Open();

                String sql = "UPDATE [Sheet8$] SET[sub]='" + x13 + "',[sub1]='" + x14 + "',[sub2]='" + x15 + "',[sub3]='" + x16 + "',[sub4]='" + x17 + "',[sub5]='" + x18 + "' WHERE [d]='3' ";
                OleDbCommand cmd = new OleDbCommand(sql, con);
                cmd.ExecuteNonQuery();
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }


            //thu
            try
            {
                OleDbConnection con = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; Data Source='D:\\course.xlsx';Extended Properties=Excel 8.0 ;");
                // OleDbDataAdapter ad = new OleDbDataAdapter();
                con.Open();

                String sql = "UPDATE [Sheet8$] SET[sub]='" + t + "',[sub1]='" + u + "',[sub2]='" + v + "',[sub3]='" + w + "',[sub4]='" + x + "',[sub5]='" + y + "' WHERE [day]='THU' ";
                OleDbCommand cmd = new OleDbCommand(sql, con);
                cmd.ExecuteNonQuery();
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            try
            {
                OleDbConnection con = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; Data Source='D:\\course.xlsx';Extended Properties=Excel 8.0 ;");
                // OleDbDataAdapter ad = new OleDbDataAdapter();
                con.Open();

                String sql = "UPDATE [Sheet8$] SET[sub]='" + x19 + "',[sub1]='" + x20 + "',[sub2]='" + x21 + "',[sub3]='" + x22 + "',[sub4]='" + x23 + "',[sub5]='" + x24 + "' WHERE [d]='4' ";
                OleDbCommand cmd = new OleDbCommand(sql, con);
                cmd.ExecuteNonQuery();
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }


            //fri
            try
            {
                OleDbConnection con = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; Data Source='D:\\course.xlsx';Extended Properties=Excel 8.0 ;");
                // OleDbDataAdapter ad = new OleDbDataAdapter();
                con.Open();

                String sql = "UPDATE [Sheet8$] SET[sub]='" + z + "',[sub1]='" + z2 + "',[sub2]='" + z3 + "',[sub3]='" + z4 + "',[sub4]='" + z5 + "',[sub5]='" + z6 + "' WHERE [day]='FRI' ";
                OleDbCommand cmd = new OleDbCommand(sql, con);
                cmd.ExecuteNonQuery();
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            try
            {
                OleDbConnection con = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; Data Source='D:\\course.xlsx';Extended Properties=Excel 8.0 ;");
                // OleDbDataAdapter ad = new OleDbDataAdapter();
                con.Open();

                String sql = "UPDATE [Sheet8$] SET[sub]='" + x25 + "',[sub1]='" + x26 + "',[sub2]='" + x27 + "',[sub3]='" + x28 + "',[sub4]='" + x29 + "',[sub5]='" + x30 + "' WHERE [d]='5' ";
                OleDbCommand cmd = new OleDbCommand(sql, con);
                cmd.ExecuteNonQuery();
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }


            //sat
            try
            {
                OleDbConnection con = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; Data Source='D:\\course.xlsx';Extended Properties=Excel 8.0 ;");
                // OleDbDataAdapter ad = new OleDbDataAdapter();
                con.Open();

                String sql = "UPDATE [Sheet8$] SET[sub]='" + z7 + "',[sub1]='" + z8 + "',[sub2]='" + z9 + "',[sub3]='" + z10 + "',[sub4]='" + z11 + "',[sub5]='" + z12 + "' WHERE [day]='SAT' ";
                OleDbCommand cmd = new OleDbCommand(sql, con);
                cmd.ExecuteNonQuery();
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            try
            {
                OleDbConnection con = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; Data Source='D:\\course.xlsx';Extended Properties=Excel 8.0 ;");
                // OleDbDataAdapter ad = new OleDbDataAdapter();
                con.Open();

                String sql = "UPDATE [Sheet8$] SET[sub]='" + x31 + "',[sub1]='" + x32 + "',[sub2]='" + x33 + "',[sub3]='" + x34 + "',[sub4]='" + x35 + "',[sub5]='" + x36 + "' WHERE [d]='6' ";
                OleDbCommand cmd = new OleDbCommand(sql, con);
                cmd.ExecuteNonQuery();
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }


            MessageBox.Show("Process Successful !");

        }

        private void button7_Click(object sender, EventArgs e)
        { //start of input data for courses
            String a = comboBox1.Text;
            String b = comboBox2.Text;
            String c = comboBox3.Text;
            String d = comboBox4.Text;
            String f = comboBox5.Text;
            String g = comboBox6.Text;

            String h = comboBox13.Text;
            String i = comboBox14.Text;
            String j = comboBox15.Text;
            String k = comboBox16.Text;
            String l = comboBox17.Text;
            String m = comboBox18.Text;

            String n = comboBox25.Text;
            String o = comboBox26.Text;
            String p = comboBox27.Text;
            String q = comboBox28.Text;
            String r = comboBox29.Text;
            String s = comboBox30.Text;

            String t = comboBox37.Text;
            String u = comboBox38.Text;
            String v = comboBox39.Text;
            String w = comboBox40.Text;
            String x = comboBox41.Text;
            String y = comboBox42.Text;

            String z = comboBox49.Text;
            String z2 = comboBox50.Text;
            String z3 = comboBox51.Text;
            String z4 = comboBox52.Text;
            String z5 = comboBox53.Text;
            String z6 = comboBox54.Text;

            String z7 = comboBox61.Text;
            String z8 = comboBox62.Text;
            String z9 = comboBox63.Text;
            String z10 = comboBox64.Text;
            String z11 = comboBox65.Text;
            String z12 = comboBox66.Text;
            //end

            //start of input data for teachers

            String x1 = comboBox7.Text;
            String x2 = comboBox8.Text;
            String x3 = comboBox9.Text;
            String x4 = comboBox10.Text;
            String x5 = comboBox11.Text;
            String x6 = comboBox12.Text;

            String x7 = comboBox19.Text;
            String x8 = comboBox20.Text;
            String x9 = comboBox21.Text;
            String x10 = comboBox22.Text;
            String x11 = comboBox23.Text;
            String x12 = comboBox24.Text;

            String x13 = comboBox31.Text;
            String x14 = comboBox32.Text;
            String x15 = comboBox33.Text;
            String x16 = comboBox34.Text;
            String x17 = comboBox35.Text;
            String x18 = comboBox36.Text;

            String x19 = comboBox43.Text;
            String x20 = comboBox44.Text;
            String x21 = comboBox45.Text;
            String x22 = comboBox46.Text;
            String x23 = comboBox47.Text;
            String x24 = comboBox48.Text;

            String x25 = comboBox55.Text;
            String x26 = comboBox56.Text;
            String x27 = comboBox57.Text;
            String x28 = comboBox58.Text;
            String x29 = comboBox59.Text;
            String x30 = comboBox60.Text;

            String x31 = comboBox67.Text;
            String x32 = comboBox68.Text;
            String x33 = comboBox69.Text;
            String x34 = comboBox70.Text;
            String x35 = comboBox71.Text;
            String x36 = comboBox72.Text;

            // end

            //mon
            try
            {
                OleDbConnection con = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; Data Source='D:\\course.xlsx';Extended Properties=Excel 8.0 ;");
                // OleDbDataAdapter ad = new OleDbDataAdapter();
                con.Open();

                String sql = "UPDATE [Sheet9$] SET[sub]='" + a + "',[sub1]='" + b + "',[sub2]='" + c + "',[sub3]='" + d + "',[sub4]='" + f + "',[sub5]='" + g + "' WHERE [day]='MON' ";
                OleDbCommand cmd = new OleDbCommand(sql, con);
                cmd.ExecuteNonQuery();
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            try
            {
                OleDbConnection con = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; Data Source='D:\\course.xlsx';Extended Properties=Excel 8.0 ;");
                // OleDbDataAdapter ad = new OleDbDataAdapter();
                con.Open();

                String sql = "UPDATE [Sheet9$] SET[sub]='" + x1 + "',[sub1]='" + x2 + "',[sub2]='" + x3 + "',[sub3]='" + x4 + "',[sub4]='" + x5 + "',[sub5]='" + x6 + "' WHERE [day]='1' ";
                OleDbCommand cmd = new OleDbCommand(sql, con);
                cmd.ExecuteNonQuery();
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            //tue
            try
            {
                OleDbConnection con = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; Data Source='D:\\course.xlsx';Extended Properties=Excel 8.0 ;");
                // OleDbDataAdapter ad = new OleDbDataAdapter();
                con.Open();

                String sql = "UPDATE [Sheet9$] SET[sub]='" + h + "',[sub1]='" + i + "',[sub2]='" + j + "',[sub3]='" + k + "',[sub4]='" + l + "',[sub5]='" + m + "' WHERE [day]='TUE' ";
                OleDbCommand cmd = new OleDbCommand(sql, con);
                cmd.ExecuteNonQuery();
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            try
            {
                OleDbConnection con = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; Data Source='D:\\course.xlsx';Extended Properties=Excel 8.0 ;");
                // OleDbDataAdapter ad = new OleDbDataAdapter();
                con.Open();

                String sql = "UPDATE [Sheet9$] SET[sub]='" + x7 + "',[sub1]='" + x8 + "',[sub2]='" + x9 + "',[sub3]='" + x10 + "',[sub4]='" + x11 + "',[sub5]='" + x12 + "' WHERE [day]='2' ";
                OleDbCommand cmd = new OleDbCommand(sql, con);
                cmd.ExecuteNonQuery();
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            //wed
            try
            {
                OleDbConnection con = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; Data Source='D:\\course.xlsx';Extended Properties=Excel 8.0 ;");
                // OleDbDataAdapter ad = new OleDbDataAdapter();
                con.Open();

                String sql = "UPDATE [Sheet9$] SET[sub]='" + n + "',[sub1]='" + o + "',[sub2]='" + p + "',[sub3]='" + q + "',[sub4]='" + r + "',[sub5]='" + s + "' WHERE [day]='WED' ";
                OleDbCommand cmd = new OleDbCommand(sql, con);
                cmd.ExecuteNonQuery();
                con.Close();
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            try
            {
                OleDbConnection con = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; Data Source='D:\\course.xlsx';Extended Properties=Excel 8.0 ;");
                // OleDbDataAdapter ad = new OleDbDataAdapter();
                con.Open();

                String sql = "UPDATE [Sheet9$] SET[sub]='" + x13 + "',[sub1]='" + x14 + "',[sub2]='" + x15 + "',[sub3]='" + x16 + "',[sub4]='" + x17 + "',[sub5]='" + x18 + "' WHERE [d]='3' ";
                OleDbCommand cmd = new OleDbCommand(sql, con);
                cmd.ExecuteNonQuery();
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }


            //thu
            try
            {
                OleDbConnection con = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; Data Source='D:\\course.xlsx';Extended Properties=Excel 8.0 ;");
                // OleDbDataAdapter ad = new OleDbDataAdapter();
                con.Open();

                String sql = "UPDATE [Sheet9$] SET[sub]='" + t + "',[sub1]='" + u + "',[sub2]='" + v + "',[sub3]='" + w + "',[sub4]='" + x + "',[sub5]='" + y + "' WHERE [day]='THU' ";
                OleDbCommand cmd = new OleDbCommand(sql, con);
                cmd.ExecuteNonQuery();
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            try
            {
                OleDbConnection con = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; Data Source='D:\\course.xlsx';Extended Properties=Excel 8.0 ;");
                // OleDbDataAdapter ad = new OleDbDataAdapter();
                con.Open();

                String sql = "UPDATE [Sheet9$] SET[sub]='" + x19 + "',[sub1]='" + x20 + "',[sub2]='" + x21 + "',[sub3]='" + x22 + "',[sub4]='" + x23 + "',[sub5]='" + x24 + "' WHERE [d]='4' ";
                OleDbCommand cmd = new OleDbCommand(sql, con);
                cmd.ExecuteNonQuery();
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }


            //fri
            try
            {
                OleDbConnection con = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; Data Source='D:\\course.xlsx';Extended Properties=Excel 8.0 ;");
                // OleDbDataAdapter ad = new OleDbDataAdapter();
                con.Open();

                String sql = "UPDATE [Sheet9$] SET[sub]='" + z + "',[sub1]='" + z2 + "',[sub2]='" + z3 + "',[sub3]='" + z4 + "',[sub4]='" + z5 + "',[sub5]='" + z6 + "' WHERE [day]='FRI' ";
                OleDbCommand cmd = new OleDbCommand(sql, con);
                cmd.ExecuteNonQuery();
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            try
            {
                OleDbConnection con = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; Data Source='D:\\course.xlsx';Extended Properties=Excel 8.0 ;");
                // OleDbDataAdapter ad = new OleDbDataAdapter();
                con.Open();

                String sql = "UPDATE [Sheet9$] SET[sub]='" + x25 + "',[sub1]='" + x26 + "',[sub2]='" + x27 + "',[sub3]='" + x28 + "',[sub4]='" + x29 + "',[sub5]='" + x30 + "' WHERE [d]='5' ";
                OleDbCommand cmd = new OleDbCommand(sql, con);
                cmd.ExecuteNonQuery();
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }


            //sat
            try
            {
                OleDbConnection con = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; Data Source='D:\\course.xlsx';Extended Properties=Excel 8.0 ;");
                // OleDbDataAdapter ad = new OleDbDataAdapter();
                con.Open();

                String sql = "UPDATE [Sheet9$] SET[sub]='" + z7 + "',[sub1]='" + z8 + "',[sub2]='" + z9 + "',[sub3]='" + z10 + "',[sub4]='" + z11 + "',[sub5]='" + z12 + "' WHERE [day]='SAT' ";
                OleDbCommand cmd = new OleDbCommand(sql, con);
                cmd.ExecuteNonQuery();
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            try
            {
                OleDbConnection con = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; Data Source='D:\\course.xlsx';Extended Properties=Excel 8.0 ;");
                // OleDbDataAdapter ad = new OleDbDataAdapter();
                con.Open();

                String sql = "UPDATE [Sheet9$] SET[sub]='" + x31 + "',[sub1]='" + x32 + "',[sub2]='" + x33 + "',[sub3]='" + x34 + "',[sub4]='" + x35 + "',[sub5]='" + x36 + "' WHERE [d]='6' ";
                OleDbCommand cmd = new OleDbCommand(sql, con);
                cmd.ExecuteNonQuery();
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }


            MessageBox.Show("Process Successful !");

        }

        private void button8_Click(object sender, EventArgs e)
        {
            //start of input data for courses
            String a = comboBox1.Text;
            String b = comboBox2.Text;
            String c = comboBox3.Text;
            String d = comboBox4.Text;
            String f = comboBox5.Text;
            String g = comboBox6.Text;

            String h = comboBox13.Text;
            String i = comboBox14.Text;
            String j = comboBox15.Text;
            String k = comboBox16.Text;
            String l = comboBox17.Text;
            String m = comboBox18.Text;

            String n = comboBox25.Text;
            String o = comboBox26.Text;
            String p = comboBox27.Text;
            String q = comboBox28.Text;
            String r = comboBox29.Text;
            String s = comboBox30.Text;

            String t = comboBox37.Text;
            String u = comboBox38.Text;
            String v = comboBox39.Text;
            String w = comboBox40.Text;
            String x = comboBox41.Text;
            String y = comboBox42.Text;

            String z = comboBox49.Text;
            String z2 = comboBox50.Text;
            String z3 = comboBox51.Text;
            String z4 = comboBox52.Text;
            String z5 = comboBox53.Text;
            String z6 = comboBox54.Text;

            String z7 = comboBox61.Text;
            String z8 = comboBox62.Text;
            String z9 = comboBox63.Text;
            String z10 = comboBox64.Text;
            String z11 = comboBox65.Text;
            String z12 = comboBox66.Text;
            //end

            //start of input data for teachers

            String x1 = comboBox7.Text;
            String x2 = comboBox8.Text;
            String x3 = comboBox9.Text;
            String x4 = comboBox10.Text;
            String x5 = comboBox11.Text;
            String x6 = comboBox12.Text;

            String x7 = comboBox19.Text;
            String x8 = comboBox20.Text;
            String x9 = comboBox21.Text;
            String x10 = comboBox22.Text;
            String x11 = comboBox23.Text;
            String x12 = comboBox24.Text;

            String x13 = comboBox31.Text;
            String x14 = comboBox32.Text;
            String x15 = comboBox33.Text;
            String x16 = comboBox34.Text;
            String x17 = comboBox35.Text;
            String x18 = comboBox36.Text;

            String x19 = comboBox43.Text;
            String x20 = comboBox44.Text;
            String x21 = comboBox45.Text;
            String x22 = comboBox46.Text;
            String x23 = comboBox47.Text;
            String x24 = comboBox48.Text;

            String x25 = comboBox55.Text;
            String x26 = comboBox56.Text;
            String x27 = comboBox57.Text;
            String x28 = comboBox58.Text;
            String x29 = comboBox59.Text;
            String x30 = comboBox60.Text;

            String x31 = comboBox67.Text;
            String x32 = comboBox68.Text;
            String x33 = comboBox69.Text;
            String x34 = comboBox70.Text;
            String x35 = comboBox71.Text;
            String x36 = comboBox72.Text;

            // end

            //mon
            try
            {
                OleDbConnection con = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; Data Source='D:\\course.xlsx';Extended Properties=Excel 8.0 ;");
                // OleDbDataAdapter ad = new OleDbDataAdapter();
                con.Open();

                String sql = "UPDATE [Sheet10$] SET[sub]='" + a + "',[sub1]='" + b + "',[sub2]='" + c + "',[sub3]='" + d + "',[sub4]='" + f + "',[sub5]='" + g + "' WHERE [day]='MON' ";
                OleDbCommand cmd = new OleDbCommand(sql, con);
                cmd.ExecuteNonQuery();
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            try
            {
                OleDbConnection con = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; Data Source='D:\\course.xlsx';Extended Properties=Excel 8.0 ;");
                // OleDbDataAdapter ad = new OleDbDataAdapter();
                con.Open();

                String sql = "UPDATE [Sheet10$] SET[sub]='" + x1 + "',[sub1]='" + x2 + "',[sub2]='" + x3 + "',[sub3]='" + x4 + "',[sub4]='" + x5 + "',[sub5]='" + x6 + "' WHERE [day]='1' ";
                OleDbCommand cmd = new OleDbCommand(sql, con);
                cmd.ExecuteNonQuery();
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            //tue
            try
            {
                OleDbConnection con = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; Data Source='D:\\course.xlsx';Extended Properties=Excel 8.0 ;");
                // OleDbDataAdapter ad = new OleDbDataAdapter();
                con.Open();

                String sql = "UPDATE [Sheet10$] SET[sub]='" + h + "',[sub1]='" + i + "',[sub2]='" + j + "',[sub3]='" + k + "',[sub4]='" + l + "',[sub5]='" + m + "' WHERE [day]='TUE' ";
                OleDbCommand cmd = new OleDbCommand(sql, con);
                cmd.ExecuteNonQuery();
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            try
            {
                OleDbConnection con = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; Data Source='D:\\course.xlsx';Extended Properties=Excel 8.0 ;");
                // OleDbDataAdapter ad = new OleDbDataAdapter();
                con.Open();

                String sql = "UPDATE [Sheet10$] SET[sub]='" + x7 + "',[sub1]='" + x8 + "',[sub2]='" + x9 + "',[sub3]='" + x10 + "',[sub4]='" + x11 + "',[sub5]='" + x12 + "' WHERE [day]='2' ";
                OleDbCommand cmd = new OleDbCommand(sql, con);
                cmd.ExecuteNonQuery();
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            //wed
            try
            {
                OleDbConnection con = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; Data Source='D:\\course.xlsx';Extended Properties=Excel 8.0 ;");
                // OleDbDataAdapter ad = new OleDbDataAdapter();
                con.Open();

                String sql = "UPDATE [Sheet10$] SET[sub]='" + n + "',[sub1]='" + o + "',[sub2]='" + p + "',[sub3]='" + q + "',[sub4]='" + r + "',[sub5]='" + s + "' WHERE [day]='WED' ";
                OleDbCommand cmd = new OleDbCommand(sql, con);
                cmd.ExecuteNonQuery();
                con.Close();
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            try
            {
                OleDbConnection con = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; Data Source='D:\\course.xlsx';Extended Properties=Excel 8.0 ;");
                // OleDbDataAdapter ad = new OleDbDataAdapter();
                con.Open();

                String sql = "UPDATE [Sheet10$] SET[sub]='" + x13 + "',[sub1]='" + x14 + "',[sub2]='" + x15 + "',[sub3]='" + x16 + "',[sub4]='" + x17 + "',[sub5]='" + x18 + "' WHERE [d]='3' ";
                OleDbCommand cmd = new OleDbCommand(sql, con);
                cmd.ExecuteNonQuery();
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }


            //thu
            try
            {
                OleDbConnection con = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; Data Source='D:\\course.xlsx';Extended Properties=Excel 8.0 ;");
                // OleDbDataAdapter ad = new OleDbDataAdapter();
                con.Open();

                String sql = "UPDATE [Sheet10$] SET[sub]='" + t + "',[sub1]='" + u + "',[sub2]='" + v + "',[sub3]='" + w + "',[sub4]='" + x + "',[sub5]='" + y + "' WHERE [day]='THU' ";
                OleDbCommand cmd = new OleDbCommand(sql, con);
                cmd.ExecuteNonQuery();
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            try
            {
                OleDbConnection con = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; Data Source='D:\\course.xlsx';Extended Properties=Excel 8.0 ;");
                // OleDbDataAdapter ad = new OleDbDataAdapter();
                con.Open();

                String sql = "UPDATE [Sheet10$] SET[sub]='" + x19 + "',[sub1]='" + x20 + "',[sub2]='" + x21 + "',[sub3]='" + x22 + "',[sub4]='" + x23 + "',[sub5]='" + x24 + "' WHERE [d]='4' ";
                OleDbCommand cmd = new OleDbCommand(sql, con);
                cmd.ExecuteNonQuery();
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }


            //fri
            try
            {
                OleDbConnection con = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; Data Source='D:\\course.xlsx';Extended Properties=Excel 8.0 ;");
                // OleDbDataAdapter ad = new OleDbDataAdapter();
                con.Open();

                String sql = "UPDATE [Sheet10$] SET[sub]='" + z + "',[sub1]='" + z2 + "',[sub2]='" + z3 + "',[sub3]='" + z4 + "',[sub4]='" + z5 + "',[sub5]='" + z6 + "' WHERE [day]='FRI' ";
                OleDbCommand cmd = new OleDbCommand(sql, con);
                cmd.ExecuteNonQuery();
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            try
            {
                OleDbConnection con = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; Data Source='D:\\course.xlsx';Extended Properties=Excel 8.0 ;");
                // OleDbDataAdapter ad = new OleDbDataAdapter();
                con.Open();

                String sql = "UPDATE [Sheet10$] SET[sub]='" + x25 + "',[sub1]='" + x26 + "',[sub2]='" + x27 + "',[sub3]='" + x28 + "',[sub4]='" + x29 + "',[sub5]='" + x30 + "' WHERE [d]='5' ";
                OleDbCommand cmd = new OleDbCommand(sql, con);
                cmd.ExecuteNonQuery();
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }


            //sat
            try
            {
                OleDbConnection con = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; Data Source='D:\\course.xlsx';Extended Properties=Excel 8.0 ;");
                // OleDbDataAdapter ad = new OleDbDataAdapter();
                con.Open();

                String sql = "UPDATE [Sheet10$] SET[sub]='" + z7 + "',[sub1]='" + z8 + "',[sub2]='" + z9 + "',[sub3]='" + z10 + "',[sub4]='" + z11 + "',[sub5]='" + z12 + "' WHERE [day]='SAT' ";
                OleDbCommand cmd = new OleDbCommand(sql, con);
                cmd.ExecuteNonQuery();
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            try
            {
                OleDbConnection con = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; Data Source='D:\\course.xlsx';Extended Properties=Excel 8.0 ;");
                // OleDbDataAdapter ad = new OleDbDataAdapter();
                con.Open();

                String sql = "UPDATE [Sheet10$] SET[sub]='" + x31 + "',[sub1]='" + x32 + "',[sub2]='" + x33 + "',[sub3]='" + x34 + "',[sub4]='" + x35 + "',[sub5]='" + x36 + "' WHERE [d]='6' ";
                OleDbCommand cmd = new OleDbCommand(sql, con);
                cmd.ExecuteNonQuery();
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }


            MessageBox.Show("Process Successful !");

        }

        private void comboBox33_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBox57_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBox72_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }
}
