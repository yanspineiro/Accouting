using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SF_Lib
{
    public partial class FormProgress : Form
    {
        public string Text_value { set; get; }
        public bool CancelProgress { set; get; }
        public FormProgress()
        {
            InitializeComponent();
        }

        private void FormProgress_Load(object sender, EventArgs e)
        {
            if (Text_value != null && Text_value != string.Empty)
            {
                Label_infoFile.Text = Text_value;
            }
            CancelProgress = false;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            CancelProgress = true;

        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
