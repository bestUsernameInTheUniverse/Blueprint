using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Reflection;
using SolidEdgeFramework;
using SolidEdgePart;


namespace Blueprint
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button_test_Click(object sender, EventArgs e)
        {
            TestClass test = new TestClass();
            test.run_test();
        }
    }
}
