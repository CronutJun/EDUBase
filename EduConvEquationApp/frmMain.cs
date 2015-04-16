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
                    lvwItem.SubItems.Add(".");
                    lvwFiles.Items.Add(lvwItem);
                }

                for (int i = 0; i < lvwFiles.Items.Count; i++)
                {
                    lvwFiles.Items[i].SubItems[1].Text = eConv.Convert(lvwFiles.Items[i].Text);
                    lvwFiles.Items[i].SubItems[2].Text = eConv.TagStr;
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
                xlSheet.Cells[i, 1] = "Image";
                xlSheet.Cells[i, 2] = item.Text;
                xlSheet.Cells[i, 3] = "'" + item.SubItems[1].Text;
                xlSheet.Cells[i, 4] = "'" + item.SubItems[2].Text;
                Excel.Shape shape = xlSheet.Shapes.AddPicture(item.Text, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, xlSheet.Cells[i, 1].left, xlSheet.Cells[i, 1].top, xlSheet.Cells[i, 1].width, xlSheet.Cells[i, 1].height);
                shape.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoCTrue;
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

        private void exportToHTMLToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string htmlText = "";
            string htmlHead = @"
<head>
<title>EDUBASE EQ EDITOR</title>
<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"" />
<script type=""text/javascript"" src=""http://cdn.mathjax.org/mathjax/2.3-latest/MathJax.js?config=TeX-AMS-MML_HTMLorMML""></script>
<script type='text/javascript' src=""http://bank.edubase.co.kr/edubank/common/js/makemathjax.js""></script>
<script type=""text/javascript"" src=""http://bank.edubase.co.kr/edubank/common/js/jquery-1.10.2.min.js""></script>
<link rel=""stylesheet"" type=""text/css"" href=""http://bank.edubase.co.kr/edubank/common/css/choproblem.asp""/>
<style>
textarea{border:0;height:100%;}
table, th, td {border: 1px solid black;border-collapse: collapse;}
th, td {
    padding: 3px !important;
}
</style>
<script type=""text/x-mathjax-config"">
  MathJax.Hub.Config({
    TeX: { Augment: {
      Definitions: {
        macros: {
          overparen: ['UnderOver','23DC'],
          underparen: ['UnderOver','23DD']
        }
      }
    }}
  });
</script>

<link rel=""stylesheet"" type=""text/css"" href=""http://bank.edubase.co.kr/edubank/common/css/eq.css""/>
<script type=""text/x-mathjax-config"">
//<![CDATA[
    MathJax.Hub.Config({extensions: [""tex2jax.js""]  });
    MathJax.Hub.Register.StartupHook(""Begin"",function (){parent.jaxload=true;});
//]]>
</script>


<script type=""text/javascript"">
//<![CDATA[
    function showOne(gid)
    {
        window.open(""m.asp?gid="" + gid);
    }

    function changeStr(s)
    {
		s=s.replace(/⇐/g,""LARROW "");
		s=s.replace(/−/g,""-"");
		s=s.replace(/⋯/g,""…"");
		s=s.replace(/⋅/g,""·"");
		s=s.replace(/∘/g,""°"");
		s=s.replace(/∅/g,""ø"");
		s=s.replace(/˚/g,""°"");
		s=s.replace(/•/g,""●"");

		s=s.replace(/〉/g,"">"");
		s = s.replace(/≈/g, ""  APPROX "");
		s = s.replace(/∉/g, "" NOTIN "");

		s=s.replace(/box{``````/g,""box{　"");
		s=s.replace(/box{````/g,""box{　"");
		s=s.replace(/box{~}/g,""box{　}"");
		s=s.replace(/box{`}/g,""box{~}"");
		s=s.replace(/box{``}/g,""box{~}"");
		s=s.replace(/box{```}/g,""box{~}"");
		s=s.replace(/box{ ` }/g,""box{~}"");
		return s;
    }
    function makeEq()
    {
        var arr=[];
        var strEqList="""";
        $('textarea').each(function(){
            strEqList+=""<EQ>""+ changeStr($(this).val()) + "" <\/EQ>""
        });

        $.ajax({
            type:""POST"",
            url:""http://bank.edubase.co.kr/edubank/common/cgi/converteq/converteq.cgi"",
            async:false,
            data:""eqlist=""+strEqList,
            success : function(data) {

                var arreqlist=data.split(""<CEQ>"");

                for(a=1;a<arreqlist.length;a++)
                {
                    strCurEq=arreqlist[a].split(""<\/CEQ>"")[0];


                    strCurEq=makeLatexString(strCurEq)
                    if (strCurEq.indexOf(""&"")>0)
                    {
                        strCurEq=strCurEq.replace(new RegExp(""&amp;"",""g""),""&"");
                    }
                    strCurEq=strCurEq.replace(new RegExp(""&lt;"",""g""),""<"");
                    strCurEq=strCurEq.replace(new RegExp(""&gt;"",""g""),"">"");

                    strCurEq=_rv2(strCurEq);
                    strCurEq=strCurEq.replace(new RegExp(""&_60;"",""g""),""<"");
                    strCurEq=strCurEq.replace(new RegExp(""&_61;"",""g""),"">"");
                    strCurEq=strCurEq.replace(new RegExp(""&nbsp;&nbsp;"",""g""),""\,"");
                    arr.push(strCurEq);
                }

            },
            error : function(xhr, status, error){
                alert(""서버에서 데이터를 읽는 중 오류가 발생하였습니다."");
            }
        });

        var index=0;
        $('textarea').each(function(){
            tid=this.id.replace(""txt"",""latex"");

            document.getElementById(tid).innerHTML=(""<textarea style='height:100%;'>""+arr[index] + ""</textarea>"");
            this.parentNode.innerHTML=(""$$"" + arr[index++] + ""$$"");
        });



        MathJax.Hub.Queue([""Typeset"",MathJax.Hub], function(){});
    }

    $(document).ready(function(){
        makeEq();
    });
//]]>
</script>
</head>
";
            string bodyTable = @"
<body style=""margin:10px;"">
<form name=""frm"" method=""post"" action=""m.asp"">

</form>
<form>
<table>
	<tr>
		<th>Index</th>
		<th>원본</th>
		<th>미리보기</th>
		<th>한글수식</th>
        <th style='display:none;'>latex변환</th>
	</tr>
";
            int i = 1;
            foreach (ListViewItem item in lvwFiles.Items)
            {
                bodyTable += "<tr align='center'>"
                           + String.Format("<td>{0}</td>", i)
                           + String.Format("<td><img src='{0}' /></td>", "../down.files/" + Path.GetFileName(item.Text))
//                           + String.Format("<td><img src='{0}' /></td>", item.Text)
                           + String.Format("<td style='font-size:11pt;'><textarea id='txt{0}'>{1}</textarea></td>", i, item.SubItems[1].Text)
                           + String.Format("<td style='font-size:11pt;'>{0}</td>", item.SubItems[1].Text)
                           + String.Format("<td style='font-size:11pt;display:none;'><div  id='latex{0}'></div></td></tr>\r\n", i);
                i++;
            }
            bodyTable += @"
</table>
</form>
</body>
";
            htmlText = @"<html>" + htmlHead + bodyTable + @"</html>";
            File.WriteAllText(String.Format("{0}\\check_mathtype_convert.html", System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)), htmlText);
  
        }
    }
}
