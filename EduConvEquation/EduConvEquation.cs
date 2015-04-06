﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace EduConvEquation
{
    public class ConvEquation
    {
        private byte[] bufMTEFHeader = { 0x21, 0xFF, 0x0B, 0x4D, 0x61, 0x74, 0x68, 0x54, 0x79, 0x70, 0x65, 0x30, 0x30, 0x31 };
        private List<byte[]> LstBlock = new List<byte[]>();
        private MTEFHeader stMTEFHeader = new MTEFHeader();
        private List<AbstractRecord> objects = new List<AbstractRecord>();
        private List<ObjectListRecord> openedObjList = new List<ObjectListRecord>();
        private AbstractRecord current = null;

        public List<AbstractRecord> Objects
        {
            get
            {
                return objects;
            }
        }

        public List<ObjectListRecord> OpenedObjList
        {
            get
            {
                return openedObjList;
            }
        }

        public string Convert(string path)
        {
            if( !File.Exists(path) )
                return "File does not exist.";
            else
            {
                FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read);

                long findPos = findMatch(fs, bufMTEFHeader);
                if (findPos >= 0)
                {
                    Console.WriteLine("Found. position = {0}", findPos);
                    CollectMTEFBlock(fs, findPos + MTEFConst.HEADER_LEN);
                    return FmtToTagStr();
                }
                else
                {
                    return "This GIF file is not a type of the MathType.";
                }
            }
        }

        public string Convert(FileStream stream)
        {
            long findPos = findMatch(stream, bufMTEFHeader);
            if (findPos >= 0)
            {
                Console.WriteLine("Found. position = {0}", findPos);
                return "Found";
            }
            else
            {
                return "Not found.";
            }
        }

        private string FmtToTagStr()
        {
            string retStr = "";
            foreach (AbstractRecord rec in objects)
            {
                if (rec.RecType == MTEFConst.REC_CHAR)
                    retStr += ((ObjectListRecord)rec).VariationStr;
                else if (rec.RecType == MTEFConst.REC_COLOR)
                    retStr += "<COLOR>";
                else if (rec.RecType == MTEFConst.REC_COLOR_DEF)
                    retStr += "<COLOR_DEF>";
                else if (rec.RecType == MTEFConst.REC_EMBELL)
                    retStr += "<EMBELL>";
                else if (rec.RecType == MTEFConst.REC_ENCODING_DEF)
                    retStr += "<ENC_DEF>";
                else if (rec.RecType == MTEFConst.REC_END)
                    retStr += "<END>";
                else if (rec.RecType == MTEFConst.REC_EQN_PREFS)
                    retStr += "<EQN_PREFS>";
                else if (rec.RecType == MTEFConst.REC_FONT_DEF)
                    retStr += "<FONT_DEF>";
                else if (rec.RecType == MTEFConst.REC_FONT_STYLE_DEF)
                    retStr += "<FONT_STYLE_DEF>";
                else if (rec.RecType == MTEFConst.REC_FUTURE)
                    retStr += "<FUTURE>";
                else if (rec.RecType == MTEFConst.REC_LINE)
                {
                    retStr += "<LINE>";
                    retStr += String.Format("[{0}]", ((ObjectListRecord)rec).ChildRecs.Count);
                }
                else if (rec.RecType == MTEFConst.REC_MATRIX)
                    retStr += "<MATRIX>";
                else if (rec.RecType == MTEFConst.REC_PILE)
                    retStr += "<PILE>";
                else if (rec.RecType == MTEFConst.REC_RULER)
                    retStr += "<RULER>";
                else if (rec.RecType == MTEFConst.REC_SIZE)
                    retStr += "<SIZE>";
                else if (rec.RecType == MTEFConst.REC_SIZE_FULL)
                    retStr += "<SIZEFULL>";
                else if (rec.RecType == MTEFConst.REC_SIZE_SUB)
                    retStr += "<SIZESUB>";
                else if (rec.RecType == MTEFConst.REC_SIZE_SUB2)
                    retStr += "<SIZESUB2>";
                else if (rec.RecType == MTEFConst.REC_SIZE_SUBSYM)
                    retStr += "<SIZESUBSYM>";
                else if (rec.RecType == MTEFConst.REC_SIZE_SYM)
                    retStr += "<SIZESYM>";
                else if (rec.RecType == MTEFConst.REC_TMPL)
                {
                    retStr += "<TMPL>";
                    retStr += String.Format("[{0}][{1:X2},{2:X2}]", ((ObjectListRecord)rec).ChildRecs.Count, ((ObjectListRecord)rec).Selector, ((ObjectListRecord)rec).Variation);
                }
            }
            return retStr;
        }

        private string FmtToHwpStr(AbstractRecord rec)
        {
            string retStr = "";

            if (rec == null) return retStr;

            if (rec.RecType == MTEFConst.REC_CHAR)
                retStr += ((ObjectListRecord)rec).VariationStr;
            else if (rec.RecType == MTEFConst.REC_LINE)
            {
                retStr += "<LINE>";
                retStr += String.Format("[{0}]", ((ObjectListRecord)rec).ChildRecs.Count);
            }
            else if (rec.RecType == MTEFConst.REC_MATRIX)
                retStr += "<MATRIX>";
            else if (rec.RecType == MTEFConst.REC_PILE)
                retStr += "<PILE>";
            else if (rec.RecType == MTEFConst.REC_RULER)
                retStr += "<RULER>";
            else if (rec.RecType == MTEFConst.REC_SIZE)
                retStr += "<SIZE>";
            else if (rec.RecType == MTEFConst.REC_SIZE_FULL)
                retStr += "<SIZEFULL>";
            else if (rec.RecType == MTEFConst.REC_SIZE_SUB)
                retStr += "<SIZESUB>";
            else if (rec.RecType == MTEFConst.REC_SIZE_SUB2)
                retStr += "<SIZESUB2>";
            else if (rec.RecType == MTEFConst.REC_SIZE_SUBSYM)
                retStr += "<SIZESUBSYM>";
            else if (rec.RecType == MTEFConst.REC_SIZE_SYM)
                retStr += "<SIZESYM>";
            else if (rec.RecType == MTEFConst.REC_TMPL)
            {
                AbstractRecord lRec = null;
                if (((ObjectListRecord)rec).Selector == 0x0B) //Fraction
                {
                    foreach (AbstractRecord crec in ((ObjectListRecord)rec).ChildRecs)
                    {
                        if (crec != ((ObjectListRecord)crec).ChildRecs.Last<AbstractRecord>())
                        {
                            retStr += FmtToHwpStr(crec) + "over";
                        }
                        else {
                            retStr += FmtToHwpStr(crec);
                        }
                        lRec = crec;
                    }
                    if (rec.NextRec != null)
                        return FmtToHwpStr(rec.NextRec);
                }
            }
            if (rec is ObjectListRecord)
            {
                if (((ObjectListRecord)rec).ChildRecs.Count > 0)
                {

                }
                else if (rec.NextRec != null)
                    return FmtToHwpStr(rec.NextRec);
            }
            else
            {
                if (rec.NextRec != null)
                    return FmtToHwpStr(rec.NextRec);
            }
            return retStr;
        }

        private long findMatch(FileStream fs, byte[] bufferToLookFor)
        {
            byte[] readBuffer = new byte[16384]; // our input buffer
            //byte[] readBuffer = new byte[100]; // our input buffer
            int bytesRead = 0; // number of bytes read
            int offset    = 0; // offset inside read-buffer
            long filePos  = 0; // position inside the file before read operation
            while( (bytesRead = fs.Read(readBuffer, offset, readBuffer.Length - offset)) > 0 )
            {
                for( int i = 0; i < bytesRead + offset - bufferToLookFor.Length; i++ )
                {
                    bool match = true;
                    for (int j = 0; j < bufferToLookFor.Length; j++)
                        if (bufferToLookFor[j] != readBuffer[i + j])
                        {
                            match = false;
                            break;
                        }
                    if (match)
                    {
                        return filePos + i - offset;
                    }
                }
                // store file position before next read
                filePos = fs.Position;

                // store the last few characters to ensure matches on "chunk boundaries"
                offset = bufferToLookFor.Length;
                for (int i = 0; i < offset; i++)
                    readBuffer[i] = readBuffer[readBuffer.Length - offset + i];
            }
            return -1;
        }

        private void CollectMTEFBlock(FileStream fs, long startPos)
        {
            fs.Seek(startPos, SeekOrigin.Begin);
            int sizeBlock    = 0;
            int totSizeBlock = 0;
            LstBlock.Clear();
            byte[] totBlock = null;
            while( (sizeBlock = fs.ReadByte()) > 0 )
            {
                Console.WriteLine("Size of Block = {0}", sizeBlock);
                byte[] blockData = new byte[sizeBlock];
                fs.Read(blockData, 0, sizeBlock);
                LstBlock.Add(blockData);
                totSizeBlock += sizeBlock;
            }
            Console.WriteLine("Count of block = {0}", LstBlock.Count);
            if (LstBlock.Count > 0)
            {                                                                                                                                                                                                                                                                                                                       
                totBlock = new byte[totSizeBlock];
                int destOffset = 0;
                foreach (byte[] elem in LstBlock)
                {
                    Console.WriteLine("elem size = {0}", elem.Length);
                    System.Buffer.BlockCopy(elem, 0, totBlock, destOffset, elem.Length);
                    destOffset += elem.Length;
                }
                // Print Hex String
                for (int i = 0; i < totBlock.Length; i++)
                    if ((i + 1) % 16 == 0)
                        Console.WriteLine("{0:X2}", totBlock[i]);
                    else
                        Console.Write("{0:X2} ", totBlock[i]);
                Console.WriteLine("");
                ParseMTEF(totBlock);
            }
        }

        private void ParseMTEF(byte[] data)
        {
            int dataPos = 0;

            // Header
            stMTEFHeader.MTEFVer       = data[0];
            stMTEFHeader.platform      = data[1];
            stMTEFHeader.productVer    = data[2];
            stMTEFHeader.productSubVer = data[3];

            dataPos = 5;
            stMTEFHeader.appKey = ByteToString(data, ref dataPos);

            dataPos++;
            stMTEFHeader.EquatOpt = data[dataPos];

            // parse MTEF Records
            Objects.Clear();
            openedObjList.Clear();
            current = null;

            dataPos++;
            parseMTEFRecords(data, dataPos);
        }

        private void parseMTEFRecords(byte[] data, int dataPos)
        {
            int dp = dataPos;

            // parse MTEF Records
            if (data[dp] == MTEFConst.REC_ENCODING_DEF)
            {
                Objects.Add(new EncodingDefRecord());
                EncodingDefRecord encDefRec = (EncodingDefRecord)Objects.Last<AbstractRecord>();
                encDefRec.PrevRec = current;
                if (current != null)
                    current.NextRec = encDefRec;
                current = encDefRec;
                encDefRec.RecType = MTEFConst.REC_ENCODING_DEF;
                dp++;
                encDefRec.EncName = ByteToString(data, ref dp);
                Console.WriteLine("dp = {0}, EncName = {1}", dp, encDefRec.EncName);
            }
            else if (data[dp] == MTEFConst.REC_FONT_DEF)
            {
                Objects.Add(new FontDefRecord());
                FontDefRecord fontDefRec = (FontDefRecord)Objects.Last<AbstractRecord>();
                fontDefRec.PrevRec = current;
                if (current != null)
                    current.NextRec = fontDefRec;
                current = fontDefRec;
                fontDefRec.RecType = MTEFConst.REC_FONT_DEF;
                dp++;
                fontDefRec.IndexOfEnc = data[dp];
                dp++;
                fontDefRec.FontName = ByteToString(data, ref dp);
                Console.WriteLine("IndexOfEnd = {0}, FontName = {1}, dataPos = {2}", fontDefRec.IndexOfEnc, fontDefRec.FontName, dp);
            }
            else if (data[dp] == MTEFConst.REC_COLOR)
            {
                dp++;
                if (data[dp] != 0x00)
                {
                    byte H = data[dp];
                    dp++;
                    byte L = data[dp];
                    dp++;
                    dp++; // 0x00
                    H = data[dp];
                    dp++;
                    L = data[dp];
                    dp++;
                }
                Console.WriteLine("Color Record Skip");
            }
            else if (data[dp] == MTEFConst.REC_COLOR_DEF)
            {
                Objects.Add(new ColorDefRecord());
                ColorDefRecord colorDefRec = (ColorDefRecord)Objects.Last<AbstractRecord>();
                colorDefRec.PrevRec = current;
                if (current != null)
                    current.NextRec = colorDefRec;
                current = colorDefRec;
                colorDefRec.RecType = data[dp];
                dp++;
                colorDefRec.Option = data[dp];
                dp++;
                colorDefRec.Color1H = data[dp];
                dp++;
                colorDefRec.Color1L = data[dp];
                dp++;
                colorDefRec.Color2H = data[dp];
                dp++;
                colorDefRec.Color2L = data[dp];
                dp++;
                colorDefRec.Color3H = data[dp];
                dp++;
                colorDefRec.Color3L = data[dp];
                if (colorDefRec.Option == MTEFConst.mtfeCOLOR_CMYK)
                {
                    dp++;
                    colorDefRec.Color4H = data[dp];
                    dp++;
                    colorDefRec.Color4L = data[dp];
                }
                colorDefRec.ColorName = ByteToString(data, ref dp);
                Console.WriteLine("ColorName = {0}, dataPos = {1}", colorDefRec.ColorName, dp);
            }
            else if (data[dp] == MTEFConst.REC_EQN_PREFS)
            {
                Objects.Add(new EqnPrefsRecord());
                EqnPrefsRecord eqnPrefsRec = (EqnPrefsRecord)Objects.Last<AbstractRecord>();
                eqnPrefsRec.PrevRec = current;
                if (current != null)
                    current.NextRec = eqnPrefsRec;
                current = eqnPrefsRec;
                eqnPrefsRec.RecType = data[dp];
                dp++;
                eqnPrefsRec.Option = data[dp];
                eqnPrefsRec.ParseNibbleData(data, ref dp);
                Console.WriteLine("dp = {0}, data = {1}", dp, data[dp]);
            }
            else if (data[dp] == MTEFConst.REC_FONT_STYLE_DEF)
            {
                Objects.Add(new FontStyleDefRecord());
                FontStyleDefRecord fontStyleDefRec = (FontStyleDefRecord)Objects.Last<AbstractRecord>();
                fontStyleDefRec.PrevRec = current;
                if (current != null)
                    current.NextRec = fontStyleDefRec;
                current = fontStyleDefRec;
                fontStyleDefRec.RecType = data[dp];
                dp++;
                fontStyleDefRec.FontDefIndex = data[dp];
                dp++;
                fontStyleDefRec.CharStyle = data[dp];
                Console.WriteLine("FontStyleDefRec.FontDefIndex = {0}, CharStyle = {1}", fontStyleDefRec.FontDefIndex, fontStyleDefRec.CharStyle);
            }
            else if (data[dp] == MTEFConst.REC_SIZE)
            {
                Objects.Add(new SizeRecord());
                SizeRecord SizeRec = (SizeRecord)Objects.Last<AbstractRecord>();
                SizeRec.PrevRec = current;
                if (current != null)
                    current.NextRec = SizeRec;
                current = SizeRec;
                SizeRec.RecType = data[dp];
                dp++;
                SizeRec.Option = data[dp];
                if (SizeRec.Option == 0x64) // 100
                {
                    dp++;
                    SizeRec.LSize = data[dp];
                    dp++;
                    SizeRec.DSize = BitConverter.ToInt16(data, dp);
                    dp++;
                }
                else if (SizeRec.Option == 0x65) // 101
                {
                    // Point size 16-bit Integer (2byte)
                    dp++;
                    SizeRec.PointSize = BitConverter.ToInt16(data, dp);
                    dp++;
                }
                else
                {
                    dp++;
                    SizeRec.LSize = data[dp];
                }
                Console.WriteLine("Size record added..{0:X2}", SizeRec.RecType);
            }
            else if (data[dp] == MTEFConst.REC_SIZE_FULL
            ||       data[dp] == MTEFConst.REC_SIZE_SUB
            ||       data[dp] == MTEFConst.REC_SIZE_SUB2
            ||       data[dp] == MTEFConst.REC_SIZE_SYM
            ||       data[dp] == MTEFConst.REC_SIZE_SUBSYM)
            {
                Objects.Add(new SizeRecord());
                SizeRecord SizeRec = (SizeRecord)Objects.Last<AbstractRecord>();
                SizeRec.PrevRec = current;
                if (current != null)
                    current.NextRec = SizeRec;
                current = SizeRec;
                SizeRec.RecType = data[dp];
                Console.WriteLine("Size record added..{0:X2}", SizeRec.RecType);
            }
            else if (data[dp] == MTEFConst.REC_LINE)
            {
                Objects.Add(new ObjectListRecord());
                ObjectListRecord objListRec = (ObjectListRecord)Objects.Last<AbstractRecord>();
                objListRec.PrevRec = current;
                if (current != null)
                    current.NextRec = objListRec;
                current = objListRec;
                objListRec.RecType = data[dp];
                dp++;
                objListRec.Option = data[dp];
                if (OpenedObjList.Count > 0)
                {
                    objListRec.ParentRec = (AbstractRecord)OpenedObjList.Last<ObjectListRecord>();
                    OpenedObjList.Last<ObjectListRecord>().ChildRecs.Add(objListRec);
                }
                if ((objListRec.Option & MTEFConst.mtfeOPT_LINE_NULL) != MTEFConst.mtfeOPT_LINE_NULL)
                    OpenedObjList.Add(objListRec);
                if ((objListRec.Option & MTEFConst.mtfeOPT_NUDGE) == MTEFConst.mtfeOPT_NUDGE)
                {
                    dp++;
                    objListRec.NudgeDX = data[dp];
                    dp++;
                    objListRec.NudgeDY = data[dp];
                }
                if ((objListRec.Option & MTEFConst.mtfeOPT_LINE_LSPACE) == MTEFConst.mtfeOPT_LINE_LSPACE)
                {
                    dp++;
                    objListRec.Selector = data[dp];
                    dp++;
                    objListRec.Variation = data[dp];
                }
                if ((objListRec.Option & MTEFConst.mtfeOPT_LP_RULER) == MTEFConst.mtfeOPT_LP_RULER)
                {
                    //dp++;
                    //objListRec.RulerRec.RecType = data[dp];
                    dp++;
                    objListRec.RulerRec.NStop = data[dp];
                    for (int i = 0; i < (int)objListRec.RulerRec.NStop; i++)
                    {
                        dp++; // tab-stop type
                        dp++; // Offset (integer)
                        objListRec.RulerRec.TabStopList = data[dp];
                        dp++;
                    }
                }
                Console.WriteLine("Line record added");
            }
            else if (data[dp] == MTEFConst.REC_TMPL)
            {
                Objects.Add(new ObjectListRecord());
                ObjectListRecord objListRec = (ObjectListRecord)Objects.Last<AbstractRecord>();
                objListRec.PrevRec = current;
                if (current != null)
                    current.NextRec = objListRec;
                current = objListRec;
                objListRec.RecType = data[dp];
                dp++;
                objListRec.Option = data[dp];
                if ((objListRec.Option & MTEFConst.mtfeOPT_NUDGE) == MTEFConst.mtfeOPT_NUDGE)
                {
                    dp++;
                    objListRec.NudgeDX = data[dp];
                    dp++;
                    objListRec.NudgeDY = data[dp];
                }
                dp++;
                objListRec.Selector = data[dp];
                dp++;
                objListRec.Variation = data[dp];
                dp++;
                objListRec.TempSpecOpt = data[dp];
                Console.WriteLine("TEMPL Record Added {0:X2}, {1:X2}, {2:X2}, {3:X2}", objListRec.Option, objListRec.Selector, objListRec.Variation, objListRec.TempSpecOpt);
                if (OpenedObjList.Count > 0)
                {
                    objListRec.ParentRec = (AbstractRecord)OpenedObjList.Last<ObjectListRecord>();
                    OpenedObjList.Last<ObjectListRecord>().ChildRecs.Add(objListRec);
                }
                OpenedObjList.Add(objListRec);
            }
            else if (data[dp] == MTEFConst.REC_CHAR)
            {
                Objects.Add(new ObjectListRecord());
                ObjectListRecord objListRec = (ObjectListRecord)Objects.Last<AbstractRecord>();
                objListRec.PrevRec = current;
                if (current != null)
                    current.NextRec = objListRec;
                current = objListRec;
                objListRec.RecType = data[dp];
                dp++;
                objListRec.Option = data[dp];
                dp++;
                // typeface
                objListRec.Selector = data[dp];
                if (objListRec.Option == 0x00)
                {
                    dp++;
                    objListRec.Variation = data[dp];
                    objListRec.VariationStr = UnicodeToString(data, dp, 2);
                    dp++; //skip lower byte
                    Console.WriteLine("Variation = {0:X2}, VariationStr = {1}", objListRec.Variation, objListRec.VariationStr);
                }
                if ((objListRec.Option & MTEFConst.mtfeOPT_NUDGE) == MTEFConst.mtfeOPT_NUDGE)
                {
                    dp++;
                    objListRec.NudgeDX = data[dp];
                    dp++;
                    objListRec.NudgeDY = data[dp];
                    if ((objListRec.Option & MTEFConst.mtfeOPT_CHAR_ENC_CHAR_8) != MTEFConst.mtfeOPT_CHAR_ENC_CHAR_8)
                    {
                        dp++;
                        objListRec.Variation = data[dp];
                        objListRec.VariationStr = UnicodeToString(data, dp, 2);
                        dp++; //skip lower byte
                        Console.WriteLine("Variation = {0:X2}, VariationStr = {1}", objListRec.Variation, objListRec.VariationStr);
                    }
                }
                if ((objListRec.Option & MTEFConst.mtfeOPT_CHAR_EMBELL) == MTEFConst.mtfeOPT_CHAR_EMBELL)
                {
                    dp++;
                    objListRec.Variation = data[dp];
                    objListRec.VariationStr = UnicodeToString(data, dp, 2);
                    dp++; //skip lower byte
                    dp++;
                    objListRec.EmbellRec.RecType = data[dp];
                    dp++;
                    objListRec.EmbellRec.Option = data[dp];
                    dp++;
                    objListRec.EmbellRec.Embell = data[dp];
                    Console.WriteLine("Variation = {0:X2}, VariationStr = {1}, Embell = {2:X2}", objListRec.Variation, objListRec.VariationStr, objListRec.EmbellRec.Embell);
                }
                if ((objListRec.Option & MTEFConst.mtfeOPT_CHAR_ENC_CHAR_8) == MTEFConst.mtfeOPT_CHAR_ENC_CHAR_8)
                {
                    dp++;
                    objListRec.VariationStr = UnicodeToString(data, dp, 2);
                    Console.WriteLine("Variation = \\u{0:X2}{1:X2}, VariationStr = {2}", data[dp + 1], data[dp], objListRec.VariationStr);
                    dp++; //Skip MTCode
                    dp++;
                    objListRec.Variation = data[dp];
                }
                if (OpenedObjList.Count > 0)
                {
                    objListRec.ParentRec = (AbstractRecord)OpenedObjList.Last<ObjectListRecord>();
                    OpenedObjList.Last<ObjectListRecord>().ChildRecs.Add(objListRec);
                }
            }
            else if (data[dp] == MTEFConst.REC_PILE)
            {
                Objects.Add(new ObjectListRecord());
                ObjectListRecord objListRec = (ObjectListRecord)Objects.Last<AbstractRecord>();
                objListRec.PrevRec = current;
                if (current != null)
                    current.NextRec = objListRec;
                current = objListRec;
                objListRec.RecType = data[dp];
                dp++;
                objListRec.Option = data[dp];
                dp++;
                objListRec.HAlign = data[dp]; //halign
                dp++;
                objListRec.VAlign = data[dp]; //valign
                if ((objListRec.Option & MTEFConst.mtfeOPT_LP_RULER) == MTEFConst.mtfeOPT_LP_RULER)
                {
                    //dp++;
                    //objListRec.RulerRec.RecType = data[dp];
                    dp++;
                    objListRec.RulerRec.NStop = data[dp];
                    for (int i = 0; i < (int)objListRec.RulerRec.NStop; i++)
                    {
                        dp++; // tab-stop type
                        dp++; // Offset (integer)
                        objListRec.RulerRec.TabStopList = data[dp];
                        dp++;
                    }
                }
                if (OpenedObjList.Count > 0)
                {
                    objListRec.ParentRec = (AbstractRecord)OpenedObjList.Last<ObjectListRecord>();
                    OpenedObjList.Last<ObjectListRecord>().ChildRecs.Add(objListRec);
                }
                OpenedObjList.Add(objListRec);
                Console.WriteLine("Pile record added");
            }
            else if (data[dp] == MTEFConst.REC_MATRIX)
            {
                Objects.Add(new ObjectListRecord());
                ObjectListRecord objListRec = (ObjectListRecord)Objects.Last<AbstractRecord>();
                objListRec.PrevRec = current;
                if (current != null)
                    current.NextRec = objListRec;
                current = objListRec;
                objListRec.RecType = data[dp];
                dp++;
                objListRec.Option = data[dp];
                dp++;
                objListRec.VAlign = data[dp]; //valign
                dp++;
                objListRec.HJust = data[dp];  //h_just
                dp++;
                objListRec.VJust = data[dp];  //v_just
                if (OpenedObjList.Count > 0)
                {
                    objListRec.ParentRec = (AbstractRecord)OpenedObjList.Last<ObjectListRecord>();
                    OpenedObjList.Last<ObjectListRecord>().ChildRecs.Add(objListRec);
                }
                OpenedObjList.Add(objListRec);
                Console.WriteLine("Matrix record added");
            }
            else if (data[dp] == MTEFConst.REC_EMBELL)
            {
                Objects.Add(new EmbellRecord());
                EmbellRecord embellRec = (EmbellRecord)Objects.Last<AbstractRecord>();
                embellRec.PrevRec = current;
                if (current != null)
                    current.NextRec = embellRec;
                current = embellRec;
                embellRec.RecType = data[dp];
                dp++;
                embellRec.Option = data[dp];
                if ((embellRec.Option & MTEFConst.mtfeOPT_NUDGE) == MTEFConst.mtfeOPT_NUDGE)
                {
                    dp++;
                    embellRec.NudgeDX = data[dp];
                    dp++;
                    embellRec.NudgeDY = data[dp];
                }
                dp++;
                embellRec.Embell = data[dp];
                if (OpenedObjList.Count > 0)
                {
                    embellRec.ParentRec = (AbstractRecord)OpenedObjList.Last<ObjectListRecord>();
                    OpenedObjList.Last<ObjectListRecord>().ChildRecs.Add(embellRec);
                }
                Console.WriteLine("Embell record added");
            }
            else if (data[dp] == MTEFConst.REC_END)
            {
                if (OpenedObjList.Count > 0)
                {
                    OpenedObjList.Remove(OpenedObjList.Last<ObjectListRecord>());
                }
            }
            else return;

            dp++;
            if (dp < data.Length)
            {
                parseMTEFRecords(data, dp);
            }
        }

        private string UnicodeToString(byte[] data, int dataPos, int length)
        {
            bool comp = false;
            comp |= CompareBytes(data, dataPos, MTEFConst.spc1);
            comp |= CompareBytes(data, dataPos, MTEFConst.spc2);
            comp |= CompareBytes(data, dataPos, MTEFConst.spc3);
            comp |= CompareBytes(data, dataPos, MTEFConst.spc4);
            if (comp)
                return " ";
            else
                return Encoding.Unicode.GetString(data, dataPos, length);
        }

        private bool CompareBytes(byte[] src, int dataPos, byte[] comp)
        {
            for (int i = 0; i < comp.Length; i++)
            {
                if (src[dataPos + i] != comp[i])
                {
                    return false;
                }
            }
            return true;
        }

        private string ByteToString(byte[] data, ref int dataPos)
        {
            byte[] strBytes = new byte[255];

            int i = 0;
            Array.Clear(strBytes, 0x00, 255);
            while (data[dataPos] != 0x00)
            {
                strBytes[i] = data[dataPos];
                dataPos++;
                i++;
            }
            strBytes[i] = 0x00;

            return Encoding.Default.GetString(strBytes, 0, i);
        }
    }
}
 