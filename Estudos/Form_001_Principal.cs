using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Estudos
{
    public partial class Form_001_Principal : Form
    {
        public Form_001_Principal()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Form_002_RichTextBox formRTB = new Form_002_RichTextBox();
            formRTB.Show();
        }
    }
}
