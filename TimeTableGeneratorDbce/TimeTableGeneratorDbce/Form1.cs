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
    public partial class main : Form
    {
        public main()
        {
            InitializeComponent();
        }

        private void course_Click(object sender, EventArgs e)
        {
            courseMenu c = new courseMenu();
            c.ShowDialog();
          

        }

        private void teacher_Click(object sender, EventArgs e)
        {
            teacherControl tc = new teacherControl();
            tc.ShowDialog();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            generateTable gt = new generateTable();
            gt.ShowDialog();
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {

        }

        private void viewTables_Click(object sender, EventArgs e)
        {
            viewTable vt = new viewTable();
            vt.ShowDialog();
        }
    }
}
