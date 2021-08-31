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
    public partial class courseMenu : Form
    {
        public courseMenu()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            addCourse ad = new addCourse();
            ad.ShowDialog();
        }

        private void button2_Click(object sender, EventArgs e)
        { 

            modifyCourse m = new modifyCourse();
            m.ShowDialog();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            courseDelete d = new courseDelete();
            d.ShowDialog();
        }
    }
}
