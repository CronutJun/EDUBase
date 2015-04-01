using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
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
            if (ofdFile.ShowDialog() == DialogResult.OK)
            {
                ConvEquation eConv = new ConvEquation();
                lblResult.Text = eConv.Convert(ofdFile.FileName);
                //textBox1.Text = Encoding.UTF8.GetString(Encoding.UTF8.GetBytes(lblResult.Text));
                textBox1.Text = lblResult.Text;
            }
        }

        private void btnCheck_Click(object sender, EventArgs e)
        {
            if (fbdDir.ShowDialog() == DialogResult.OK)
            {
                ConvEquation eConv = new ConvEquation();
                string[] paths = Directory.GetFiles(@fbdDir.SelectedPath, "*.gif");
                foreach (string str in paths)
                {
                    ListViewItem lvwItem = new ListViewItem();
                    lvwItem.Text = str;
                    lvwItem.SubItems.Add(".");
                    lvwFiles.Items.Add(lvwItem);
                }

                for (int i = 0; i < lvwFiles.Items.Count; i++)
                {
                    lvwFiles.Items[i].SubItems[1].Text = eConv.Convert(lvwFiles.Items[i].Text);
                    Application.DoEvents();
                }
                MessageBox.Show("Finish");
            }
        }

        private void lvwFiles_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (lvwFiles.SelectedItems.Count > 0)
                pictureBox1.BackgroundImage = Image.FromFile(@lvwFiles.SelectedItems[0].Text);
        }
    }
}
