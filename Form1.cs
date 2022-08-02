using JARVISCS;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace JarvisFnl
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
           
            
        }

        private void timer1_Tick(object sender, EventArgs e)
        {

            panel2.Width += 3;
            if (panel2.Width >= 682)
            {
                timer1.Stop();
                MainForm fm2 = new MainForm();
                fm2.Show();
                this.Hide();
            }

        }
    }
}
