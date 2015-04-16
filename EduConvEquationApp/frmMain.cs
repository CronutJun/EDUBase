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
        s=s.replace(/?/g,""LARROW "");
        s=s.replace(/?/g,""-"");
        s=s.replace(/?/g,""…"");
        s=s.replace(/?/g,""·"");
        s=s.replace(/?/g,""°"");
        s=s.replace(/?/g,""ø"");
        s=s.replace(/˚/g,""°"");
        s=s.replace(/?/g,""●"");

        s=s.replace(/?/g,"">"");
        s = s.replace(/?/g, ""  APPROX "");
        s = s.replace(/?/g, "" NOTIN "");

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
    function _rv2(x, blat)
    {
        x=x.replace(new RegExp(""&_60;"",""g""),""<"");
        x=x.replace(new RegExp(""&_61;"",""g""),"">"");
        x=x.replace(new RegExp(""_#91;"",""g""),""["");
        x=x.replace(new RegExp(""&_93;"",""g""),""]"");
        x=x.replace(new RegExp(""&_123;"",""g""),""{"");
        x=x.replace(new RegExp(""&_125;"",""g""),""}"");
        x=x.replace(new RegExp(""&_149;"",""g""),""?"");
        x=x.replace(new RegExp(""曹曹"",""g""),""\\,"");		//　
        x=x.replace(new RegExp(""曹"",""g""),""\\,"");

        x=x.replace(new RegExp(String.fromCharCode(160),""g""),""~"");

        x=x.replace(new RegExp(""ŋ"",""g""),""ㆍ"");
        x=x.replace(new RegExp(""&_40;"",""g""),""("");
        x=x.replace(new RegExp(""&_41;"",""g""),"")"");

        x=x.replace(new RegExp(""%"",""g""),""\\&#x0025;"");



        if(blat==null)blat=false;
        if(blat==false)
        {
            x=x.replace(new RegExp(""空白"",""g""),"""");
        }
        return x;
    }



    function makeLatexString(s)
    {
        s=s.replace(new RegExp(""&_60;I&_61;"",""g""),""<I>"");
        s=s.replace(new RegExp(""&_123;"",""g""),""\\&_123;"");
        s=s.replace(new RegExp(""&_125;"",""g""),""\\&_125;"");
        s=s.replace(new RegExp(""_SPLIT_"",""g""),""?"");
        s=s.replace(new RegExp(""~"",""g""),unescape(""%u200B"")+""~"");

        var x=s.indexOf("":"");

        var stxt= _rv2(convertEqFormat(s.substring(x+1),false));

        stxt=stxt.replace(/_mpara\(/g,""["");
        stxt=stxt.replace(/_mpara\)/g,""]"");
        stxt=stxt.replace(new RegExp(""?"",""g""),""-"");
        stxt=stxt.replace(/`SPACE`/g,""~"");
        stxt=stxt.replace(new RegExp(""空白"",""g""),"" "");
        stxt=stxt.replace(/\s\s/g,""　~"");
        stxt=stxt.replace(new RegExp(String.fromCharCode(160)+String.fromCharCode(160),""g""),""　~"");
        stxt=makeBox2Input(stxt);

        stxt=stxt.replace(new RegExp(String.fromCharCode(160),""g""),"""");
        stxt=stxt.replace(new RegExp(unescape(""%uFEFF""),""g""),"""");
        stxt=stxt.replace(/〉/g,""&gt;"");
        stxt=stxt.replace(/〈/g,""&lt;"");
        stxt=stxt.replace(/＞/g,""&gt;"");
        stxt=stxt.replace(/＜/g,""&lt;"");
        stxt=stxt.replace(new RegExp(""?"",""g""),""ㆍ"");
        stxt=stxt.replace(/○/g,""\\bigcirc"");
        stxt=stxt.replace(/…/g,""\\cdots"");
        stxt=stxt.replace(/?/g,""\\cdot"");

        return ""\\mathrm{"" + stxt + ""}"";

    }

    function makeBox2Input(stxt)
    {
        var arr=stxt.split(""\\bbox"");
        for(var i=1;i<arr.length;i++)
        {
            var c=0;
            var colindex=0;
            for(j=0;j<arr[i].length;j++)
            {

                if(arr[i].substring(j,j+1)==""{"")c++;
                if(arr[i].substring(j,j+1)==""}"")c--;
                if(c==0)
                {
                    var strtmp=arr[i].substring(0,j+1);
                    for(k=0;k<15;k++)
                    {
                        if(strtmp.indexOf(""①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮"".substring(k,k+1))>=0)
                        {
                            colindex=k+1;
                            strtmp=strtmp.replace(""①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮"".substring(k,k+1),""　~"");
                        }
                    }
                    if(colindex>0)
                    {
                        var strchk=strtmp;
                        strchk=strchk.replace(/{/g,"""");
                        strchk=strchk.replace(/}/g,"""");
                        strchk=strchk.replace(/~/g,"""");
                        strchk=strchk.replace(/　/g,"""");
                        strchk=strchk.replace(/\s/g,"""");
                        strchk=strchk.replace(new RegExp(""\\\\mathrm"",""g""),"""");
                        strchk=strchk.replace(new RegExp(String.fromCharCode(160),""g""),"""");
                        strchk=strchk.replace(new RegExp(String.fromCharCode(8203),""g""),"""");

                        if(strchk=="""")
                        {
                            arr[i]=strtmp+arr[i].substring(j+1);
                        }
                        else
                            colindex=0;
                    }
                    break;
                }
            }
            if(colindex>0)
            {
                var h=(+colindex).toString(16);
                arr[i]=""[2px,border:1px solid #00000""+h+"";]""+arr[i];
            }
            else
                arr[i]=""[2px,border:1px solid #000000;]""+arr[i];
        }
        return arr.join(""\\bbox"");
    }


    function convertEqFormat(strTxt,bml)
    {
        var index=0;
        var arrList=[];
        var strHead, strTail, strBody;
        var strTmp;
        while(true)
        {
            var i=strTxt.indexOf(""["");
            if(i==-1)
            {
                return convertEFormatOne(strTxt,bml);
            }
            var icount=1;
            strTmp="""";
            for(var j=i+1;j<strTxt.length;j++)
            {
                icount+=(strTxt.substring(j,j+1)==""["");
                icount-=(strTxt.substring(j,j+1)==""]"");
                if(icount==0)break;
                strTmp+=strTxt.substring(j,j+1);
            }
            strHead=strTxt.substring(0,i)
            strTail=strTxt.substring(j+1);
            strBody=convertEqFormat(strTmp,bml);
            strTxt=strHead + strBody + strTail;
        }
        strTxt=_rv2(strTxt,true);
        return strTxt;
    }

    function IsNumeric(v)
    {
        return (v - 0) == v && (v+'').replace(/^\s+|\s+$/g, """").length > 0;
    }

    function convertEFormatOne(strTxt,bml)
    {
        var bmsg=false;
        if(strTxt.substring(0,3)==""<I>"")
        {

            if(bml)
            {
                return makeString2MathML(strTxt.replace(/[\{,\}]/g, """").substring(3));
            }
            else
                return ""\\mathit{"" + strTxt.replace(/[\{,\}]/g, """").substring(3) + ""}"";
        }
        else if(strTxt.substring(0,1)==""("" && strTxt.substring(1,2)>=""a"" && strTxt.substring(1,2)<=""z"")
        {
            var x=strTxt.indexOf("")"");
            var skind=strTxt.substring(1,x);
            strTxt=strTxt.substring(x+2);
            strTxt=strTxt.substring(0,strTxt.length-1);
            var arrCol=strTxt.split(""?"")
            arrCol.push("""");arrCol.push("""");arrCol.push("""");
            if(bml==false)
            {
                for(var i=0;i<arrCol.length;i++)
                {
                    if(IsNumeric(arrCol[i])==false)
                    {
                        if(arrCol[i].substring(0,1)!=""\\"")
                        {
                            arrCol[i]=""\\mathrm{""+arrCol[i]+""}"";
                        }
                    }
                }
            }
            switch(skind.toLowerCase())
            {
                case ""over"":
                    if(bml)
                        strTxt=""\\frac{""+ arrCol[0] + ""}{"" + arrCol[1] + ""}"";
                    else
                        strTxt=""\\frac{""+ arrCol[0] + ""}{"" + arrCol[1] + ""}"";
                    break;
                case ""box"":
                    strTxt=""\\bbox{"" + arrCol[0] + ""}"";
                    break;
                case ""sqrt"":
                    strTxt=""\\sqrt{"" + arrCol[0] + ""}"";
                    break;
                case ""sqrt2"":
                    strTxt=""\\sqrt_mpara(""+arrCol[0]+""_mpara){"" + arrCol[1] + ""}"";
                    break;
                case ""supsub"":
                    strTxt=""{""+arrCol[0]+""}_{"" + arrCol[2] + ""}^{"" + arrCol[1] + ""}"";
                    break;
                case ""sub"":
                    strTxt= ""{""+arrCol[0]+""}"" + ""_{"" + arrCol[1] + ""}"";
                    break;
                case ""sup"":
                    strTxt= ""{""+arrCol[0]+""}"" + ""^{"" + arrCol[1] + ""}"";
                    break;
                case ""bar"":
                    strTxt=""\\overline{""+arrCol[0]+""}"";
                    break;
                case ""vec"":
                    strTxt=""\\overrightarrow{""+arrCol[0]+""}"";
                    break;
                case ""dyad"":
                    strTxt=""\\overleftrightarrow{""+arrCol[0]+""}"";
                    break;
                case ""under"":
                    strTxt=""\\underline{""+arrCol[0]+""}"";
                    break;
                case ""dot"":
                case ""check"":
                case ""acute"":
                case ""grave"":
                    strTxt=""\\""+skind.toLowerCase()+""{""+arrCol[0]+""}"";
                    break;
                case ""hat"":
                    strTxt=""\\widehat{""+arrCol[0]+""}"";
                case ""tilde"":
                    strTxt=""\\widetilde{""+arrCol[0]+""}"";
                    break;
                case ""arch"":
                    strTxt=""\\overset{\\mmlToken{mo}{&#x23DC;}}{""+arrCol[0]+""}"";
                    break;
                case ""atop"":

                    strTxt=""\\begin{matrix}{""+arrCol[0]+""}\\\\{""+arrCol[1]+""}\\end{matrix}~"";
                    break;
                case ""over2"":
                    strTxt=""\\begin{matrix}{""+arrCol[0]+""}&{}&{}\\\\{}&{/}&{}\\\\{}&{}&{""+arrCol[1]+""}\\end{matrix}""
                    if(bmsg)alert(""존재하지 않는 latex문법 - matrix으로 대체.."");
                    break;
                case ""over3"":
                    strTxt=""\\begin{matrix}{""+arrCol[0]+""}&{/}&{""+arrCol[1]+""}\\end{matrix}"";
                    break;

                case ""rle"":
                    strTxt=""\\mathrel{\\substack{""+arrCol[0]+""\\\\\\longleftrightarrow\\\\""+arrCol[1]+""}} ""
                    break;
                case ""rel"":
                    strTxt=""\\xrightarrow_mpara(""+arrCol[0]+""_mpara){""+arrCol[1]+""}"";
                    if(bmsg)alert(""존재하지 않는 latex문법 - 비슷한 것으로 대체.."");
                    break;
                case ""lel"":
                    strTxt=""\\xleftarrow_mpara(""+arrCol[0]+""){""+arrCol[1]+""}"";
                    if(bmsg)alert(""존재하지 않는 latex문법 - 비슷한 것으로 대체.."");
                    break;
                case""integral_1"":
                    strTxt=""{\\int}"";
                    break;
                case""integral_2"":
                    strTxt=""\\int_{""+arrCol[1]+""}^{""+arrCol[0]+""}"";
                    break;
                case""integral2_1"":
                    strTxt=""{\\iint}"";
                    break;
                case""integral2_2"":
                    strTxt=""\\iint_{""+arrCol[1]+""}^{""+arrCol[0]+""}"";
                    break;
                case""integral3_1"":
                    strTxt=""{\\iiint}"";
                    break;
                case""integral3_2"":
                    strTxt=""\\iiint_{""+arrCol[1]+""}^{""+arrCol[0]+""}"";
                    break;
                case""ointegral_1"":
                    strTxt=""{\\oint}"";
                    break;
                case""ointegral_2"":
                    strTxt=""\\oint_{""+arrCol[1]+""}^{""+arrCol[0]+""}"";
                    break;
                case""ointegral2_1"":
                    strTxt=""{\\oiint}"";
                    if(bmsg)alert(""지원되지 않음"");
                    break;
                case""ointegral2_2"":
                    strTxt=""\\oiint_{""+arrCol[1]+""}^{""+arrCol[0]+""}"";
                    if(bmsg)alert(""지원되지 않음"");
                    break;
                case""ointegral3_1"":
                    strTxt=""{\\oiiint}"";
                    if(bmsg)alert(""지원되지 않음"");
                    break;
                case""ointegral3_2"":
                    strTxt=""\\oiiint_{""+arrCol[1]+""}^{""+arrCol[0]+""}"";
                    if(bmsg)alert(""지원되지 않음"");
                    break;
                case ""lim"":
                case ""max"":
                case ""min"":
                    strTxt=""\\""+skind.toLowerCase()+""_{""+arrCol[1]+""}{""+arrCol[0]+""}"";
                    break;
                case ""log"":
                case ""ln"":
                case ""sin"":
                case ""cos"":
                case ""tan"":
                case ""csc"":
                case ""sec"":
                case ""cot"":
                case ""arcsin"":
                case ""arccos"":
                case ""arctan"":
                    strTxt=""\\mathrm{""+skind.toLowerCase()+""}{""+arrCol[0]+""}"";
                    break;
                case ""log2"":
                    strTxt=""\\mathrm{""+skind.toLowerCase().replace(""2"","""")+""}_{""+arrCol[0]+""}{""+arrCol[1]+""}"";
                    break;
                case ""sin2"":
                case ""cos2"":
                case ""tan2"":
                case ""csc2"":
                case ""sec2"":
                case ""cot2"":
                case ""arcsin2"":
                case ""arccos2"":
                case ""arctan2"":
                    strTxt=""\\mathrm{""+skind.toLowerCase().replace(""2"","""")+""}^{""+arrCol[0]+""}{""+arrCol[1]+""}"";
                    break;
                case ""sum"":
                case ""prod"":
                case ""coprod"":
                case ""bigwedge"":
                case ""bigvee"":
                case ""bigoplus"":
                case ""bigodot"":
                    strTxt=""\\""+skind.toLowerCase()+""_{""+arrCol[1]+""}^{""+arrCol[0]+""}"";
                    break;
                case ""union"":
                    strTxt=""\\bigcup_{""+arrCol[1]+""}^{""+arrCol[0]+""}"";
                    break;
                case ""inter"":
                    strTxt=""\\bigcap_{""+arrCol[1]+""}^{""+arrCol[0]+""}"";
                    break;
                case ""bigotime"":
                case ""bigominus"":
                case ""bigodiv"":
                    if(bmsg)alert(""latex에서 지원하지 않는 수식"");
                    break;
                case ""not"":
                    for(var xx=0;xx<2;xx++)
                    {
                        if(arrCol[0].substring(0,8)==""\\mathrm{""  && arrCol[0].substring(arrCol[0].length-1,arrCol[0].length-0)==""}"")
                        {
                            arrCol[0]=arrCol[0].substring(8);
                            arrCol[0]=arrCol[0].substring(0,arrCol[0].length-1);
                        }
                        if(arrCol[0].substring(0,8)==""\\mathit{""  && arrCol[0].substring(arrCol[0].length-1,arrCol[0].length-0)==""}"")
                        {
                            arrCol[0]=arrCol[0].substring(8);
                            arrCol[0]=arrCol[0].substring(0,arrCol[0].length-1);
                        }
                    }
                    strTxt=""\\not""+arrCol[0]+"""";
                    break;
                case ""abs"":
                    strTxt=""\\left|{""+arrCol[0]+""}\\right|"";
                    break;
                case ""abs2"":
                    strTxt=""\\left_mpara({""+arrCol[0]+""}\\right_mpara)"";
                    break;
                case ""para1"":
                    strTxt=""\\left\\{{""+arrCol[0]+""}\\right\\}"";
                    break;
                case ""para2"":
                    strTxt=""\\left({""+arrCol[0]+""}\\right)"";
                    break;
                case ""para3"":
                    strTxt=""\\left<{""+arrCol[0]+""}\\right>"";
                    break;




                case ""rabs"":
                    strTxt=""\\left.{""+arrCol[0]+""}\\right|"";
                    break;
                case ""rabs2"":
                    strTxt=""\\left.{""+arrCol[0]+""}\\right_mpara)"";
                    break;
                case ""rpara1"":
                    strTxt=""\\left.{""+arrCol[0]+""}\\right\\}"";
                    break;
                case ""rpara2"":
                    strTxt=""\\left.{""+arrCol[0]+""}\\right)"";
                    break;
                case ""rpara3"":
                    strTxt=""\\left.{""+arrCol[0]+""}\\right>"";
                    break;

                case ""labs"":
                    strTxt=""\\left|{""+arrCol[0]+""}\\right."";
                    break;
                case ""labs2"":
                    strTxt=""\\left_mpara({""+arrCol[0]+""}\\right."";
                    break;
                case ""lpara1"":
                    strTxt=""\\left\\{{""+arrCol[0]+""}\\right."";
                    break;
                case ""lpara2"":
                    strTxt=""\\left({""+arrCol[0]+""}\\right."";
                    break;
                case ""lpara3"":
                    strTxt=""\\left<{""+arrCol[0]+""}\\right."";
                    break;


                case ""rle"":
                    if(bmsg)alert(""지원하지 않는 수식"");
                    break;
                default:
                    if(skind.toLowerCase().indexOf(""matrix_"")>=0)
                    {
                        var arrm=skind.toLowerCase().split(""_"");
                        var row=parseInt(arrm[1],10),col=parseInt(arrm[2],10);
                        var arrtable=[];

                        var ord=0;
                        for(var i=0;i<row;i++)
                        {
                            arrtable[i]=[];
                            for(var j=0;j<col;j++)
                            {
                                if(ord>arrCol.length-1)
                                    arrtable[i][j]="""";
                                else
                                    arrtable[i][j]=arrCol[ord];
                                ord++;
                            }
                        }
                        var matrixflag=arrm[0];
                        if(arrm[0]==""dmatrix"")
                            matrixflag=""vmatrix"";
                        else if(arrm[0]==""cmatrix"")
                        {
                            matrixflag=""cases"";
                        }
                        else if(arrm[0]==""rmatrix"")
                            matrixflag=""array}{rr"";
                        else if(arrm[0]==""lmatrix"")
                            matrixflag=""array}{ll"";

                        strTxt="""";

                        strTxt+=""\\begin{""+matrixflag+""}"";
                        for(var i=0;i<row;i++)
                        {
                            for(var j=0;j<col;j++)
                            {
                                strTxt+=""{""+arrtable[i][j]+""}"";
                                if(j<col-1)strTxt += ""&"";
                            }
                            if(i<row-1)strTxt+=""\\\\"";
                        }
                        if(arrm[0]==""rmatrix"" || arrm[0]==""lmatrix"")
                            strTxt+=""\\end{array}"";
                        else
                            strTxt+=""\\end{""+matrixflag+""}"";

                        strTxt=""\\!"" + strTxt + ""\\!"";
                    }
            }
        }
        return strTxt;
    }


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
        <th style='display:none;'>latex변환</th>	</tr>
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
