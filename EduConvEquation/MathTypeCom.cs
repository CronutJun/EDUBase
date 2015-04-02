using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EduConvEquation
{
    public struct MTEFHeader
    {
        private int _n;

        public byte   MTEFVer;
        public byte   platform;
        public byte   product;
        public byte   productVer;
        public byte   productSubVer;
        public String appKey;
        public byte   EquatOpt;

        public override bool Equals(object obj)
        {
            MTEFHeader target = (MTEFHeader)obj;
            return _n == target._n;                
        }

        public override int GetHashCode()
        {
            return _n.GetHashCode();
        }
    }

    public struct MTEFConst
    {
        public const int HEADER_LEN = 14;

        // Records Type
        public const byte REC_END            = 0x00;
        public const byte REC_LINE           = 0x01;
        public const byte REC_CHAR           = 0x02;
        public const byte REC_TMPL           = 0x03;
        public const byte REC_PILE           = 0x04;
        public const byte REC_MATRIX         = 0x05;
        public const byte REC_EMBELL         = 0x06;
        public const byte REC_RULER          = 0x07;
        public const byte REC_FONT_STYLE_DEF = 0x08;
        public const byte REC_SIZE           = 0x09;
        public const byte REC_SIZE_FULL      = 0x0A;
        public const byte REC_SIZE_SUB       = 0x0B;
        public const byte REC_SIZE_SUB2      = 0x0C;
        public const byte REC_SIZE_SYM       = 0x0D;
        public const byte REC_SIZE_SUBSYM    = 0x0E;
        public const byte REC_COLOR          = 0x0F;
        public const byte REC_COLOR_DEF      = 0x10;
        public const byte REC_FONT_DEF       = 0x11;
        public const byte REC_EQN_PREFS      = 0x12;
        public const byte REC_ENCODING_DEF   = 0x13;
        public const byte REC_FUTURE         = 0x64;

        // Record Option
        public const byte mtfeOPT_NUDGE = 0x08;
        public const byte mtfeOPT_CHAR_EMBELL        = 0x01;
        public const byte mtfeOPT_CHAR_FUNC_START    = 0x02;
        public const byte mtfeOPT_CHAR_ENC_CHAR_8    = 0x04;
        public const byte mtfeOPT_CHAR_ENC_CHAR_16   = 0x10;
        public const byte mtfeOPT_CHAR_ENC_NO_MTCODE = 0x20;
        public const byte mtfeOPT_LINE_NULL          = 0x01;
        public const byte mtfeOPT_LINE_LSPACE        = 0x04;
        public const byte mtfeOPT_LP_RULER           = 0x02;

        public const byte mtfeCOLOR_CMYK             = 0x01;
        public const byte mtfeCOLOR_SPOT             = 0x02;
        public const byte mtfeCOLOR_NAME             = 0x04;

        // Force space convert
        public static readonly byte[] spc1 = new byte[] { 0x02, 0xEF };
        public static readonly byte[] spc2 = new byte[] { 0x05, 0x00 };
        public static readonly byte[] spc3 = new byte[] { 0x00, 0xF7 };
        public static readonly byte[] spc4 = new byte[] { 0x08, 0xEF };

    }

    public abstract class AbstractRecord
    {
        private byte recType;
        private byte option;

        public byte RecType
        {
            get
            {
                return recType;
            }
            set
            {
                recType = value;
            }
        }

        public byte Option
        {
            get
            {
                return option;
            }
            set
            {
                option = value;
            }
        }
    }

    public class EncodingDefRecord : AbstractRecord
    {
        private string encName;

        public string EncName
        {
            get
            {
                return encName;
            }
            set
            {
                encName = value;
            }
        }
    }

    public class ColorDefRecord : AbstractRecord
    {
        private byte color1H;
        private byte color1L;
        private byte color2H;
        private byte color2L;
        private byte color3H;
        private byte color3L;
        private byte color4H;
        private byte color4L;
        private string colorName;

        public byte Color1H
        {
            get
            {
                return color1H;
            }
            set
            {
                color1H = value;
            }
        }
        public byte Color1L
        {
            get
            {
                return color1L;
            }
            set
            {
                color1L = value;
            }
        }
        public byte Color2H
        {
            get
            {
                return color2H;
            }
            set
            {
                color2H = value;
            }
        }
        public byte Color2L
        {
            get
            {
                return color2L;
            }
            set
            {
                color2L = value;
            }
        }
        public byte Color3H
        {
            get
            {
                return color3H;
            }
            set
            {
                color3H = value;
            }
        }
        public byte Color3L
        {
            get
            {
                return color3L;
            }
            set
            {
                color3L = value;
            }
        }
        public byte Color4H
        {
            get
            {
                return color4H;
            }
            set
            {
                color4H = value;
            }
        }
        public byte Color4L
        {
            get
            {
                return color4L;
            }
            set
            {
                color4L = value;
            }
        }
        public string ColorName
        {
            get
            {
                return colorName;
            }
            set
            {
                colorName = value;
            }
        }
    }

    public class FontDefRecord : AbstractRecord
    {
        private byte indexOfEnc;
        private string fontName;

        public byte IndexOfEnc
        {
            get
            {
                return indexOfEnc;
            }
            set
            {
                indexOfEnc = value;
            }
        }

        public string FontName
        {
            get
            {
                return fontName;
            }
            set
            {
                fontName = value;
            }
        }
    }

    public class RulerRecord : AbstractRecord
    {
        private byte nStop;
        private byte tabStopList;

        public byte NStop
        {
            get
            {
                return nStop;
            }
            set
            {
                nStop = value;
            }
        }
        public byte TabStopList
        {
            get
            {
                return tabStopList;
            }
            set
            {
                tabStopList = value;
            }
        }
    }

    public class FontStyleDefRecord : AbstractRecord
    {
        private byte fontDefIndex;
        private byte charStyle;

        public byte FontDefIndex
        {
            get
            {
                return fontDefIndex;
            }
            set
            {
                fontDefIndex = value;
            }
        }
        public byte CharStyle
        {
            get
            {
                return charStyle;
            }
            set
            {
                charStyle = value;
            }
        }
    }

    public class DimArray
    {
        private string size;
        private int unit;

        public string Size
        {
            get
            {
                return size;
            }
            set
            {
                size = value;
            }
        }

        public int Unit
        {
            get
            {
                return unit;
            }
            set
            {
                unit = value;
            }
        }
    }

    public class EqnPrefsRecord : AbstractRecord
    {
        private int sizeCnt  = 0;
        private int spaceCnt = 0;
        private int styleCnt = 0;

        private List<DimArray> sizeArr = new List<DimArray>();
        private List<DimArray> spaceArr = new List<DimArray>();
        private List<DimArray> styleArr = new List<DimArray>();

        public int SizeCnt {
            get {
                return sizeCnt;
            }
            set {
                sizeCnt = value;
                sizeArr.Clear();
                for( int i=0; i < sizeCnt; i++ ) {
                    sizeArr.Add(new DimArray());
                }
            }
        }
        public List<DimArray> SizeArr {
            get {
                return sizeArr;
            }
        }

        public int SpaceCnt {
            get {
                return spaceCnt;
            }
            set {
                spaceCnt = value;
                spaceArr.Clear();
                for( int i=0; i < spaceCnt; i++ ) {
                    spaceArr.Add(new DimArray());
                }
            }
        }
        public List<DimArray> SpaceArr {
            get {
                return spaceArr;
            }
        }

        public int StyleCnt {
            get {
                return styleCnt;
            }
            set {
                styleCnt = value;
                styleArr.Clear();
                for( int i=0; i < styleCnt; i++ ) {
                    styleArr.Add(new DimArray());
                }
            }
        }
        public List<DimArray> StyleArr {
            get {
                return styleArr;
            }
        }

        public void ParseNibbleData(byte[] data, ref int dataPos)
        {
            int totCnt = 0;
            int upperNibble = 0;
            int lowerNibble = 0;

            dataPos++;
            SizeCnt = data[dataPos];
            dataPos++;

            bool unitYN = true;
            while (totCnt < sizeCnt)
            {
                upperNibble = data[dataPos] & 0xF0;
                if( upperNibble == 0xF0 )
                {
                    totCnt++;
                    unitYN = true;
                }
                else {
                    makeNibbleData(0, upperNibble, unitYN, totCnt);
                    unitYN = false;
                }

                lowerNibble = (data[dataPos] & 0x0F) << 4;
                if( lowerNibble == 0xF0 )
                {
                    totCnt++;
                    unitYN = true;
                }
                else {
                    if (upperNibble == 0xF0 && lowerNibble == 0x00) { unitYN = false; dataPos++; break; }
                    makeNibbleData(0, lowerNibble, unitYN, totCnt);
                    unitYN = false;
                }

                dataPos++;
            }
            Console.WriteLine("Size Cnt = {0}", totCnt);
            foreach (DimArray size in SizeArr)
            {
                Console.WriteLine("Size = {0}", size.Size);
            }

            SpaceCnt = data[dataPos];
            dataPos++;
            totCnt = 0;
            unitYN = true;
            while (totCnt < spaceCnt)
            {
                upperNibble = data[dataPos] & 0xF0;
                if (upperNibble == 0xF0)
                {
                    totCnt++;
                    unitYN = true;
                }
                else
                {
                    makeNibbleData(1, upperNibble, unitYN, totCnt);
                    unitYN = false;
                }

                lowerNibble = (data[dataPos] & 0x0F) << 4;
                if (lowerNibble == 0xF0)
                {
                    totCnt++;
                    unitYN = true;
                }
                else
                {
                    if (upperNibble == 0xF0 && lowerNibble == 0x00) { unitYN = false; dataPos++; break; }
                    makeNibbleData(1, lowerNibble, unitYN, totCnt);
                    unitYN = false;
                }

                dataPos++;
            }
            Console.WriteLine("Space Cnt = {0}", totCnt);
            foreach (DimArray size in SpaceArr)
            {
                Console.WriteLine("Space = {0}", size.Size);
            }

            StyleCnt = data[dataPos];
            dataPos++;
            totCnt = 0;
            unitYN = false;
            while (totCnt < styleCnt)
            {
                if (data[dataPos] == 0x00)
                    break;

                //Console.WriteLine("{0:X2} {1:X2}", data[dataPos], data[dataPos + 1]);
                StyleArr[totCnt].Size = data[dataPos].ToString();
                dataPos++;

                StyleArr[totCnt].Unit = data[dataPos];
                dataPos++;

                totCnt++;
            }
            Console.WriteLine("Style Cnt = {0}, Tot Cnt = {1}", styleCnt, totCnt);
            if (styleCnt == totCnt) dataPos--;
            foreach (DimArray size in StyleArr)
            {
                Console.WriteLine("Style = {0}, Unit = {1}", size.Size, size.Unit);
            }
        }

        private void makeNibbleData(int type, int nibbleData, bool unitYN, int idx) 
        {
            if( type == 0 )
            {
                if( unitYN )
                {
                    SizeArr[idx].Unit = nibbleData;
                }
                else
                {
                    if (nibbleData == 0x00)
                    {
                        SizeArr[idx].Size +="0";
                    }
                    else if (nibbleData == 0x10)
                    {
                        SizeArr[idx].Size += "1";
                    }
                    else if (nibbleData == 0x20)
                    {
                        SizeArr[idx].Size += "2";
                    }
                    else if (nibbleData == 0x30)
                    {
                        SizeArr[idx].Size += "3";
                    }
                    else if (nibbleData == 0x40)
                    {
                        SizeArr[idx].Size += "4";
                    }
                    else if (nibbleData == 0x50)
                    {
                        SizeArr[idx].Size += "5";
                    }
                    else if (nibbleData == 0x60)
                    {
                        SizeArr[idx].Size += "6";
                    }
                    else if (nibbleData == 0x70)
                    {
                        SizeArr[idx].Size += "7";
                    }
                    else if (nibbleData == 0x80)
                    {
                        SizeArr[idx].Size += "8";
                    }
                    else if (nibbleData == 0x90)
                    {
                        SizeArr[idx].Size += "9";
                    }
                    else if (nibbleData == 0xA0)
                    {
                        SizeArr[idx].Size += ".";
                    }
                    else if (nibbleData == 0xB0)
                    {
                        SizeArr[idx].Size += "-";
                    }
                }
            }
            else if (type == 1)
            {
                if (unitYN)
                {
                    SpaceArr[idx].Unit = nibbleData;
                }
                else
                {
                    if (nibbleData == 0x00)
                    {
                        SpaceArr[idx].Size += "0";
                    }
                    else if (nibbleData == 0x10)
                    {
                        SpaceArr[idx].Size += "1";
                    }
                    else if (nibbleData == 0x20)
                    {
                        SpaceArr[idx].Size += "2";
                    }
                    else if (nibbleData == 0x30)
                    {
                        SpaceArr[idx].Size += "3";
                    }
                    else if (nibbleData == 0x40)
                    {
                        SpaceArr[idx].Size += "4";
                    }
                    else if (nibbleData == 0x50)
                    {
                        SpaceArr[idx].Size += "5";
                    }
                    else if (nibbleData == 0x60)
                    {
                        SpaceArr[idx].Size += "6";
                    }
                    else if (nibbleData == 0x70)
                    {
                        SpaceArr[idx].Size += "7";
                    }
                    else if (nibbleData == 0x80)
                    {
                        SpaceArr[idx].Size += "8";
                    }
                    else if (nibbleData == 0x90)
                    {
                        SpaceArr[idx].Size += "9";
                    }
                    else if (nibbleData == 0xA0)
                    {
                        SpaceArr[idx].Size += ".";
                    }
                    else if (nibbleData == 0xB0)
                    {
                        SpaceArr[idx].Size += "-";
                    }
                }
            }
        }
    }

    public class EmbellRecord : AbstractRecord
    {
        private byte embell;
        public byte Embell
        {
            get
            {
                return embell;
            }
            set
            {
                embell = value;
            }
        }
    }

    public class ObjectListRecord : AbstractRecord
    {
        private AbstractRecord prevRec = null;
        private AbstractRecord nextRec = null;
        private AbstractRecord parentRec = null;
        private AbstractRecord childRec = null;
        private byte selector;
        private byte nudgeDX;
        private byte nudgeDY;
        private byte variation;
        private byte tempSpecOpt;
        private string variationStr;
        private byte halign;
        private byte valign;
        private byte hjust;
        private byte vjust;
        private EmbellRecord embellRec = new EmbellRecord();
        private RulerRecord rulerRec = new RulerRecord();

        public AbstractRecord PrevRec
        {
            get
            {
                return prevRec;
            }
            set
            {
                prevRec = value;
            }
        }
        public AbstractRecord NextRec
        {
            get
            {
                return nextRec;
            }
            set
            {
                nextRec = value;
            }
        }
        public AbstractRecord ParentRec
        {
            get
            {
                return parentRec;
            }
            set
            {
                parentRec = value;
            }
        }
        public AbstractRecord ChildRec
        {
            get
            {
                return childRec;
            }
            set
            {
                childRec = value;
            }
        }
        public byte Selector
        {
            get
            {
                return selector;
            }
            set
            {
                selector = value;
            }
        }
        public byte NudgeDX
        {
            get
            {
                return nudgeDX;
            }
            set
            {
                nudgeDX = value;
            }
        }
        public byte NudgeDY
        {
            get
            {
                return nudgeDY;
            }
            set
            {
                nudgeDY = value;
            }
        }
        public byte Variation
        {
            get
            {
                return variation;
            }
            set
            {
                variation = value;
            }
        }
        public byte TempSpecOpt
        {
            get
            {
                return tempSpecOpt;
            }
            set
            {
                tempSpecOpt = value;
            }
        }
        public string VariationStr
        {
            get
            {
                return variationStr;
            }
            set
            {
                variationStr = value;
            }
        }
        public byte HAlign
        {
            get
            {
                return halign;
            }
            set
            {
                halign = value;
            }
        }
        public byte VAlign
        {
            get
            {
                return valign;
            }
            set
            {
                valign = value;
            }
        }
        public byte HJust
        {
            get
            {
                return hjust;
            }
            set
            {
                hjust = value;
            }
        }
        public byte VJust
        {
            get
            {
                return vjust;
            }
            set
            {
                vjust = value;
            }
        }
        public EmbellRecord EmbellRec
        {
            get
            {
                return embellRec;
            }
        }
        public RulerRecord RulerRec
        {
            get
            {
                return rulerRec;
            }
        }
    }

    public class EndRecord : AbstractRecord
    {

    }
}
