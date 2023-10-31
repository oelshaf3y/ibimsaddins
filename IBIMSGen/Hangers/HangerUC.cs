using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace IBIMSGen.Hangers
{
    public partial class HangerUC : UserControl
    {
        public CheckedListBox worksetnames{get; set;}
        public HangerUC()
        {
            InitializeComponent();
            this.worksetnames = worksetNames;
        }

    }
}
