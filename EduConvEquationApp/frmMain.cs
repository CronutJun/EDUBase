using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using EduConvEquation;

namespace EduConvEquationApp
{
    public partial class frmMain : Form
    {
        public frmMain()
        {
            InitializeComponent();
        }

        private void btnCall_Click(object sender, EventArgs e)
        {
            ofdFile.ShowDialog();
            ConvEquation eConv = new ConvEquation();
            lblResult.Text = eConv.Convert(ofdFile.FileName);
        }
    }
}
