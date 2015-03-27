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

    }

    public abstract class AbstractRecord
    {
        private byte recType;
        private byte option;
        private List<AbstractRecord> objects = new List<AbstractRecord>();

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

        public List<AbstractRecord> Objects
        {
            get
            {
                return objects;
            }
        }
    }

    public class EndRecord : AbstractRecord {

    }
}
