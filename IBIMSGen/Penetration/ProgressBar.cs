using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace IBIMSGen
{
    public partial class ProgressBar : Form
    {
        public int ipb = 0;
        public ProgressBar()
        {
            InitializeComponent();
        }

        private void Progress_Bar_Load(object sender, EventArgs e)
        {
            progresBarRatio.Value = ipb;

        }
    }
}
