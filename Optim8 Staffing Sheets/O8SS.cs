using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Optim8_Staffing_Sheets
{
    public partial class O8SS : Form
    {
        public O8SS()
        {
            InitializeComponent();
        }

        private void rdSSbtn_Click(object sender, EventArgs e)
        {
            Program.parkServices = false;
        }

        private void psSSbtn_Click(object sender, EventArgs e)
        {
            Program.parkServices = true;
        }
    }
}
