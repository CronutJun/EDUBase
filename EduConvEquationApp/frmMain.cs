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
using Excel = Microsoft.Office.Interop.Excel;

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

        private void lvwFiles_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                if( lvwFiles.Items.Count > 0 )
                    contextMenuStrip1.Show(Cursor.Position);
            }
        }

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            Excel.Application xlApp;
            Excel.Workbook xlBook;
            Excel.Worksheet xlSheet;

            object misValue = System.Reflection.Missing.Value;

            xlApp = new Excel.Application();
            xlBook = xlApp.Workbooks.Add(misValue);

            xlSheet = (Excel.Worksheet)xlBook.Worksheets.get_Item(1);

            int i = 1;
            foreach (ListViewItem item in lvwFiles.Items)
            {
                xlSheet.Cells[i, 1] = item.Text;
                xlSheet.Cells[i, 2] = "'" + item.SubItems[1].Text;
                i++;
            }

            xlBook.SaveAs(String.Format("{0}\\csharp-Excel.xls", System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)), Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlBook.Close(true, misValue, misValue);
            xlApp.Quit();

            releaseObject(xlSheet);
            releaseObject(xlBook);
            releaseObject(xlApp);
        }

        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Exception Occured while releasing object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}
