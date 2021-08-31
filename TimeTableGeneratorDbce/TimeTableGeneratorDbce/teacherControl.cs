using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TimeTableGeneratorDbce
{
    public partial class teacherControl : Form
    {
        public teacherControl()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            teacherAdd ta = new teacherAdd();
            ta.ShowDialog();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            modifyTeacher mt = new modifyTeacher();
            mt.ShowDialog();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            deleteTeacher dt = new deleteTeacher();
            dt.ShowDialog();
        }
    }
}
