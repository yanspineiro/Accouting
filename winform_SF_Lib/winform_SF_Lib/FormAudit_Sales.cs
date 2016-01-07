using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace winform_SF_Lib
    {
    public partial class FormAudit_Sales : Form
        {
        public DateTime From { set; get; }
        public DateTime To { set; get; }
        public bool isclose { set; get; }
        public FormAudit_Sales()
            {
            InitializeComponent();
            }

        private void button1_Click(object sender, EventArgs e)
            {
            isclose = true;
            this.Close();
            }

        private void button8_Click(object sender, EventArgs e)
            {
            From = dtp_From.Value;
            To = dtp_to.Value;
            isclose = false;
            this.Close();
            }
        }
    }
