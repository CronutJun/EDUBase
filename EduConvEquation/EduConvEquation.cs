using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Text.RegularExpressions;

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
        private bool hasMatrix = false;
        private byte alignOpt = 0x00;
        private string tagStr = "";
        private String fontName = "";

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
            if (!File.Exists(path))
                return "File does not exist.";
            else
            {
                FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read);

                return Convert(fs);
            }
        }

        public string Convert(FileStream stream)
        {
            long findPos = findMatch(stream, bufMTEFHeader);
            if (findPos >= 0)
            {
                Console.WriteLine("Found. position = {0}", findPos);
                CollectMTEFBlock(stream, findPos + MTEFConst.HEADER_LEN);
                return ReplaceString(FmtToHwpStr(objects.Count > 0 ? objects[0] : null, true, false, 0x00, 0x00, 0));
            }
            else
            {
                tagStr = "This GIF file is not a type of the MathType.";
                return tagStr;
            }
        }

        private string ReplaceString(string input)
        {
            string replaced = input;
            if (alignOpt == 0x04)
            {
                char[] d = { '#' };
                string[] arr = replaced.Split(d);
                for (int i = 0; i < arr.Length; i++)
                {
                    Regex r = new Regex("([=<>])");
                    arr[i] = r.Replace(arr[i], "&$1", 1);
                }
                replaced = String.Join("#", arr);
            }
            replaced = replaced.Replace("<−", "<`−")
                               .Replace("−>", "−`>")
                               .Replace("<=", "<`=")
                               .Replace("=>", "=`>");
            return replaced;
        }

        public string TagStr
        {
            get
            {
                if (!tagStr.StartsWith("This GIF file is not"))
                    return FmtToTagStr();
                else
                    return tagStr;
            }
        }

        private string AdjustPile(ObjectListRecord parent, AbstractRecord current, string formated)
        {
            if (parent.RecType == MTEFConst.REC_PILE
            || parent.RecType == MTEFConst.REC_LINE)
            {
                if (parent.LimitCount > 0 && parent.ChildRecs.Count != parent.LimitCount)
                {
                    if (current.RecType == MTEFConst.REC_TMPL && ((ObjectListRecord)current).Selector == 0x17)
                        return formated;
                    else
                    {
                        return String.Format(" rpile{{~#{0}}}", formated);
                    }
                }
                else return formated;
            }
            else return formated;
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
                {
                    retStr += "<MATRIX>";
                    retStr += String.Format("[{0}]", ((ObjectListRecord)rec).ChildRecs.Count);
                }
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

        private string FmtToHwpStr(AbstractRecord rec, bool keepNextRecord, bool skipBrace, byte selector, byte variation, int limitType)
        {
            string retStr = "", trimStr = "";

            if (rec == null) return retStr;

            if (rec.RecType == MTEFConst.REC_CHAR)
            {
                if (!((ObjectListRecord)rec).VariationStr.Equals("{") && !((ObjectListRecord)rec).VariationStr.Equals("}"))
                {
                    retStr = ((ObjectListRecord)rec).VariationStr;
                    trimStr = retStr.Replace("`", "").Replace("~", "");
                    if ((((ObjectListRecord)rec).Option & MTEFConst.mtfeOPT_CHAR_EMBELL) == MTEFConst.mtfeOPT_CHAR_EMBELL)
                    {
                        if (((ObjectListRecord)rec).EmbellRec.Embell == 0x0A
                        || ((ObjectListRecord)rec).EmbellRec.Embell == 0x16)
                            retStr = trimStr.Length > 0 ? "{not " + retStr + "}" : retStr;
                    }
                    if (((ObjectListRecord)rec).Selector == 0x81
                    || ((ObjectListRecord)rec).Selector == 0x82)
                        retStr = trimStr.Length > 0 ? "{rm " + retStr + "}" : retStr;
                    if (trimStr.Length > 0)
                        if (limitType == 1)
                            retStr = " rpile{~#" + retStr + "}";
                        else if (limitType == 2)
                            retStr = " rpile{" + retStr + "#~}";
                }
            }
            else if (rec.RecType == MTEFConst.REC_COLOR)
                retStr += "";
            else if (rec.RecType == MTEFConst.REC_COLOR_DEF)
                retStr += "";
            else if (rec.RecType == MTEFConst.REC_EMBELL)
                retStr += "";
            else if (rec.RecType == MTEFConst.REC_ENCODING_DEF)
                retStr += "";
            else if (rec.RecType == MTEFConst.REC_END)
                retStr += "";
            else if (rec.RecType == MTEFConst.REC_EQN_PREFS)
                retStr += "";
            else if (rec.RecType == MTEFConst.REC_FONT_DEF)
                retStr += "";
            else if (rec.RecType == MTEFConst.REC_FONT_STYLE_DEF)
                retStr += "";
            else if (rec.RecType == MTEFConst.REC_FUTURE)
                retStr += "";
            else if (rec.RecType == MTEFConst.REC_LINE && ((ObjectListRecord)rec).ChildRecs.Count > 0)
            {
                int lType = 0;
                if (((ObjectListRecord)rec).LimitCount > 0 && ((ObjectListRecord)rec).ChildRecs.Count != ((ObjectListRecord)rec).LimitCount)
                {
                    foreach (AbstractRecord crec in ((ObjectListRecord)rec).ChildRecs)
                    {
                        if (crec.RecType == MTEFConst.REC_TMPL && ((ObjectListRecord)crec).Selector == 0x17 && ((ObjectListRecord)crec).ChildRecs.Count > 1)
                        {
                            if (((ObjectListRecord)((ObjectListRecord)crec).ChildRecs[1]).ChildRecs.Count == 0)
                                lType = 1;
                            else lType = 2;
                            break;
                        }
                    }
                }
                if (skipBrace)
                    retStr += FmtToHwpStr(((ObjectListRecord)rec).ChildRecs[0], true, false, selector, variation, lType);
                else
                    retStr += "{" + FmtToHwpStr(((ObjectListRecord)rec).ChildRecs[0], true, false, selector, variation, lType) + "}";
                //retStr += " " + FmtToHwpStr(((ObjectListRecord)rec).ChildRecs[0], true, false, selector, variation) + " ";
            }
            else if (rec.RecType == MTEFConst.REC_MATRIX)
            {
                int i = 0;
                foreach (AbstractRecord crec in ((ObjectListRecord)rec).ChildRecs)
                {
                    if (i > 0)
                        retStr += "#" + FmtToHwpStr(crec, false, false, selector, variation, 0);
                    else
                        if (selector == 0x02 && variation == 0x01)
                            retStr += " cases{" + FmtToHwpStr(crec, false, false, 0x00, 0x00, 0);
                        else
                            retStr += " matrix{" + FmtToHwpStr(crec, false, false, selector, variation, 0);
                    i++;
                }
                if (i > 0)
                    retStr += "}";
            }
            else if (rec.RecType == MTEFConst.REC_PILE)
            {
                int i = 0;
                foreach (AbstractRecord crec in ((ObjectListRecord)rec).ChildRecs)
                {
                    if (i > 0)
                        if (selector == 0x02 && variation == 0x01)
                            retStr += "#" + FmtToHwpStr(crec, false, false, selector, variation, 0);
                        else
                            retStr += "#" + FmtToHwpStr(crec, false, false, selector, variation, 0);
                    else
                        if (selector == 0x02 && variation == 0x01)
                            retStr += " cases{" + FmtToHwpStr(crec, false, false, selector, variation, 0);
                        else
                        {
                            if (((ObjectListRecord)rec).HAlign == 0x01)
                                retStr += " lpile{" + FmtToHwpStr(crec, false, false, selector, variation, 0);
                            else if (((ObjectListRecord)rec).HAlign == 0x02)
                                retStr += " pile{" + FmtToHwpStr(crec, false, false, selector, variation, 0);
                            else if (((ObjectListRecord)rec).HAlign == 0x03)
                                retStr += " rpile{" + FmtToHwpStr(crec, false, false, selector, variation, 0);
                            else
                                retStr += " lpile{" + FmtToHwpStr(crec, false, false, selector, variation, 0);
                            alignOpt = ((ObjectListRecord)rec).HAlign;
                        }
                    i++;
                }
                if (i > 0)
                    retStr += "}";
            }
            else if (rec.RecType == MTEFConst.REC_RULER)
                retStr += "";
            else if (rec.RecType == MTEFConst.REC_SIZE)
                retStr += "";
            else if (rec.RecType == MTEFConst.REC_SIZE_FULL)
                retStr += "";
            else if (rec.RecType == MTEFConst.REC_SIZE_SUB)
                retStr += "";
            else if (rec.RecType == MTEFConst.REC_SIZE_SUB2)
                retStr += "";
            else if (rec.RecType == MTEFConst.REC_SIZE_SUBSYM)
                retStr += "";
            else if (rec.RecType == MTEFConst.REC_SIZE_SYM)
                retStr += "";
            else if (rec.RecType == MTEFConst.REC_TMPL)
            {
                if (((ObjectListRecord)rec).Selector == 0x01) //Parentheses
                {
                    if ((((ObjectListRecord)rec).Variation & 0x03) == 0x03)
                        retStr += " left" + FmtToHwpStr(((ObjectListRecord)rec).ChildRecs[1], false, false, ((ObjectListRecord)rec).Selector, ((ObjectListRecord)rec).Variation, 0)
                               + FmtToHwpStr(((ObjectListRecord)rec).ChildRecs[0], true, false, ((ObjectListRecord)rec).Selector, ((ObjectListRecord)rec).Variation, 0)
                               + " right" + FmtToHwpStr(((ObjectListRecord)rec).ChildRecs[2], false, false, ((ObjectListRecord)rec).Selector, ((ObjectListRecord)rec).Variation, 0);
                    else if ((((ObjectListRecord)rec).Variation & 0x01) == 0x01)
                        retStr += " left" + FmtToHwpStr(((ObjectListRecord)rec).ChildRecs[1], false, false, ((ObjectListRecord)rec).Selector, ((ObjectListRecord)rec).Variation, 0)
                               + FmtToHwpStr(((ObjectListRecord)rec).ChildRecs[0], true, false, ((ObjectListRecord)rec).Selector, ((ObjectListRecord)rec).Variation, 0);
                    else if ((((ObjectListRecord)rec).Variation & 0x02) == 0x02)
                        retStr += FmtToHwpStr(((ObjectListRecord)rec).ChildRecs[0], true, false, ((ObjectListRecord)rec).Selector, ((ObjectListRecord)rec).Variation, 0)
                               + " right" + FmtToHwpStr(((ObjectListRecord)rec).ChildRecs[1], false, false, ((ObjectListRecord)rec).Selector, ((ObjectListRecord)rec).Variation, 0);
                }
                else if (((ObjectListRecord)rec).Selector == 0x02) //Braces
                {
                    if (hasMatrix && ((ObjectListRecord)rec).Variation == 0x01)
                    //if (hasMatrix)
                    {
                        if ((((ObjectListRecord)rec).Variation & 0x03) == 0x03)
                            retStr += FmtToHwpStr(((ObjectListRecord)rec).ChildRecs[0], true, false, ((ObjectListRecord)rec).Selector, ((ObjectListRecord)rec).Variation, 0);
                        else if ((((ObjectListRecord)rec).Variation & 0x01) == 0x01)
                            retStr += FmtToHwpStr(((ObjectListRecord)rec).ChildRecs[0], true, false, ((ObjectListRecord)rec).Selector, ((ObjectListRecord)rec).Variation, 0);
                        else if ((((ObjectListRecord)rec).Variation & 0x02) == 0x02)
                            retStr += FmtToHwpStr(((ObjectListRecord)rec).ChildRecs[0], true, false, ((ObjectListRecord)rec).Selector, ((ObjectListRecord)rec).Variation, 0);
                    }
                    else
                    {
                        if (((ObjectListRecord)rec).Variation == 0x03)
                            retStr += " left{" + FmtToHwpStr(((ObjectListRecord)rec).ChildRecs[1], false, false, ((ObjectListRecord)rec).Selector, ((ObjectListRecord)rec).Variation, 0)
                                   + FmtToHwpStr(((ObjectListRecord)rec).ChildRecs[0], true, false, ((ObjectListRecord)rec).Selector, ((ObjectListRecord)rec).Variation, 0)
                                   + " right}" + FmtToHwpStr(((ObjectListRecord)rec).ChildRecs[2], false, false, ((ObjectListRecord)rec).Selector, ((ObjectListRecord)rec).Variation, 0);
                        else if (((ObjectListRecord)rec).Variation == 0x01)
                            retStr += " left{" + FmtToHwpStr(((ObjectListRecord)rec).ChildRecs[1], false, false, ((ObjectListRecord)rec).Selector, ((ObjectListRecord)rec).Variation, 0)
                                   + FmtToHwpStr(((ObjectListRecord)rec).ChildRecs[0], true, false, ((ObjectListRecord)rec).Selector, ((ObjectListRecord)rec).Variation, 0);
                        else if (((ObjectListRecord)rec).Variation == 0x02)
                            retStr += FmtToHwpStr(((ObjectListRecord)rec).ChildRecs[0], true, false, ((ObjectListRecord)rec).Selector, ((ObjectListRecord)rec).Variation, 0)
                                   + " right}" + FmtToHwpStr(((ObjectListRecord)rec).ChildRecs[1], false, false, ((ObjectListRecord)rec).Selector, ((ObjectListRecord)rec).Variation, 0);
                    }
                }
                else if (((ObjectListRecord)rec).Selector == 0x04) //Fences
                {
                        if (((ObjectListRecord)rec).Variation == 0x03)
                            retStr += " left|" 
                                   + FmtToHwpStr(((ObjectListRecord)rec).ChildRecs[0], true, false, ((ObjectListRecord)rec).Selector, ((ObjectListRecord)rec).Variation, 0)
                                   + " right|";
                        else if (((ObjectListRecord)rec).Variation == 0x01)
                            retStr += " left|"
                                   + FmtToHwpStr(((ObjectListRecord)rec).ChildRecs[0], true, false, ((ObjectListRecord)rec).Selector, ((ObjectListRecord)rec).Variation, 0);
                        else if (((ObjectListRecord)rec).Variation == 0x02)
                            retStr += FmtToHwpStr(((ObjectListRecord)rec).ChildRecs[0], true, false, ((ObjectListRecord)rec).Selector, ((ObjectListRecord)rec).Variation, 0)
                                   + " right|"; 
                }
                else if (((ObjectListRecord)rec).Selector == 0x0A) //Root(SQRT)
                {
                    retStr += " SQRT{" + FmtToHwpStr(((ObjectListRecord)rec).ChildRecs[0], true, false, ((ObjectListRecord)rec).Selector, ((ObjectListRecord)rec).Variation, 0) + "}";
                }
                else if (((ObjectListRecord)rec).Selector == 0x0B) //Fraction
                {
                    foreach (AbstractRecord crec in ((ObjectListRecord)rec).ChildRecs)
                    {
                        if (!ReferenceEquals(crec, ((ObjectListRecord)rec).ChildRecs.Last<AbstractRecord>()))
                        {
                            retStr += FmtToHwpStr(crec, true, false, ((ObjectListRecord)rec).Selector, ((ObjectListRecord)rec).Variation, 0) + "over";
                        }
                        else
                        {
                            retStr += FmtToHwpStr(crec, true, false, ((ObjectListRecord)rec).Selector, ((ObjectListRecord)rec).Variation, 0);
                        }
                    }
                }
                else if (((ObjectListRecord)rec).Selector == 0x0C) //Under bar
                {
                    retStr += " under{";
                    foreach (AbstractRecord crec in ((ObjectListRecord)rec).ChildRecs)
                    {
                        retStr += FmtToHwpStr(crec, true, false, ((ObjectListRecord)rec).Selector, ((ObjectListRecord)rec).Variation, 0);
                    }
                    retStr += "}";
                }
                else if (((ObjectListRecord)rec).Selector == 0x0D) //bar
                {
                    retStr += " bar{";
                    foreach (AbstractRecord crec in ((ObjectListRecord)rec).ChildRecs)
                    {
                        retStr += FmtToHwpStr(crec, true, false, ((ObjectListRecord)rec).Selector, ((ObjectListRecord)rec).Variation, 0);
                    }
                    retStr += "}";
                }
                else if (((ObjectListRecord)rec).Selector == 0x17) //Limit
                {
                    if (((ObjectListRecord)rec).ChildRecs.Count <= 1)
                    {

                    }
                    else
                    {
                        if (((ObjectListRecord)((ObjectListRecord)rec).ChildRecs[1]).ChildRecs.Count == 0)
                        {
                            retStr += " rpile{";
                            retStr += "`" + FmtToHwpStr(((ObjectListRecord)rec).ChildRecs[2], true, false, ((ObjectListRecord)rec).Selector, ((ObjectListRecord)rec).Variation, 0);
                            retStr += "#`" + FmtToHwpStr(((ObjectListRecord)rec).ChildRecs[0], true, false, ((ObjectListRecord)rec).Selector, ((ObjectListRecord)rec).Variation, 0);
                            retStr += "}";
                        }
                        else
                        {
                            retStr += " rpile{";
                            retStr += "`" + FmtToHwpStr(((ObjectListRecord)rec).ChildRecs[0], true, false, ((ObjectListRecord)rec).Selector, ((ObjectListRecord)rec).Variation, 0);
                            retStr += "#`" + FmtToHwpStr(((ObjectListRecord)rec).ChildRecs[1], true, false, ((ObjectListRecord)rec).Selector, ((ObjectListRecord)rec).Variation, 0);
                            retStr += "}";
                        }
                    }
                }
                else if (((ObjectListRecord)rec).Selector == 0x1C) //Superscript
                {
                    retStr += " ^{" + FmtToHwpStr(((ObjectListRecord)rec).ChildRecs[0], true, false, ((ObjectListRecord)rec).Selector, ((ObjectListRecord)rec).Variation, 0) + "}";
                }
                else if (((ObjectListRecord)rec).Selector == 0x24) //Strike
                {
                    if ((((ObjectListRecord)rec).Variation & 0x03) == 0x02) // Not
                    {
                        retStr += "{not " + FmtToHwpStr(((ObjectListRecord)rec).ChildRecs[0], true, true, ((ObjectListRecord)rec).Selector, ((ObjectListRecord)rec).Variation, 0) + "}";
                    }
                }
                else if (((ObjectListRecord)rec).Selector == 0x25) //Box
                {
                    retStr += " box{";
                    foreach (AbstractRecord crec in ((ObjectListRecord)rec).ChildRecs)
                    {
                        retStr += FmtToHwpStr(crec, true, false, ((ObjectListRecord)rec).Selector, ((ObjectListRecord)rec).Variation, 0);
                    }
                    retStr += "}";
                }
            }
            if (rec is ObjectListRecord)
            {
                if (((ObjectListRecord)rec).ChildRecs.Count > 0)
                {
                    if ((rec.RecType == MTEFConst.REC_TMPL
                       || rec.RecType == MTEFConst.REC_PILE
                       || rec.RecType == MTEFConst.REC_MATRIX) && keepNextRecord)
                        retStr += FmtToHwpStr(rec.NextRec, true, false, selector, variation, limitType);
                }
                else if (keepNextRecord)
                    retStr += FmtToHwpStr(rec.NextRec, true, true, selector, variation, limitType);
            }
            else
            {
                if (keepNextRecord)
                    retStr += FmtToHwpStr(rec.NextRec, true, true, selector, variation, limitType);
            }
            return retStr;
        }

        private long findMatch(FileStream fs, byte[] bufferToLookFor)
        {
            byte[] readBuffer = new byte[16384]; // our input buffer
            //byte[] readBuffer = new byte[100]; // our input buffer
            int bytesRead = 0; // number of bytes read
            int offset = 0; // offset inside read-buffer
            long filePos = 0; // position inside the file before read operation
            while ((bytesRead = fs.Read(readBuffer, offset, readBuffer.Length - offset)) > 0)
            {
                for (int i = 0; i < bytesRead + offset - bufferToLookFor.Length; i++)
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
            int sizeBlock = 0;
            int totSizeBlock = 0;
            LstBlock.Clear();
            byte[] totBlock = null;
            while ((sizeBlock = fs.ReadByte()) > 0)
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
            stMTEFHeader.MTEFVer = data[0];
            stMTEFHeader.platform = data[1];
            stMTEFHeader.productVer = data[2];
            stMTEFHeader.productSubVer = data[3];

            dataPos = 5;
            stMTEFHeader.appKey = ByteToString(data, ref dataPos);

            dataPos++;
            stMTEFHeader.EquatOpt = data[dataPos];

            // parse MTEF Records
            Objects.Clear();
            openedObjList.Clear();
            current = null;
            hasMatrix = false;
            alignOpt = 0x00;
            tagStr = "";
            fontName = "";

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
                fontName = fontDefRec.FontName;
                Console.WriteLine("IndexOfEnd = {0}, FontName = {1}, dataPos = {2}", fontDefRec.IndexOfEnc, fontDefRec.FontName, dp);
            }
            else if (data[dp] == MTEFConst.REC_COLOR)
            {
                dp++;
                if (data[dp] != 0x00) // on
                {
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
                dp++;
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
            || data[dp] == MTEFConst.REC_SIZE_SUB
            || data[dp] == MTEFConst.REC_SIZE_SUB2
            || data[dp] == MTEFConst.REC_SIZE_SYM
            || data[dp] == MTEFConst.REC_SIZE_SUBSYM)
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
                {
                    OpenedObjList.Add(objListRec);
                    current = null;
                }
                if ((objListRec.Option & MTEFConst.mtfeOPT_NUDGE) == MTEFConst.mtfeOPT_NUDGE)
                {
                    dp++;
                    objListRec.NudgeDX = data[dp];
                    dp++;
                    objListRec.NudgeDY = data[dp];
                    if (objListRec.NudgeDX == 0x80 && objListRec.NudgeDY == 0x80)
                    {
                        dp++;
                        dp++;
                        dp++;
                        dp++;
                    }
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
                    if (objListRec.NudgeDX == 0x80 && objListRec.NudgeDY == 0x80)
                    {
                        dp++;
                        dp++;
                        dp++;
                        dp++;
                    }
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
                    ObjectListRecord lastPtr = OpenedObjList.Last<ObjectListRecord>();
                    OpenedObjList.Last<ObjectListRecord>().ChildRecs.Add(objListRec);
                    if (objListRec.Selector == 0x17) //Limits
                    {
                        lastPtr.LimitCount++;
                        Console.WriteLine("ChildRec count = {0}, LimitCount = {1}", lastPtr.ChildRecs.Count, lastPtr.LimitCount);
                    }
                }
                OpenedObjList.Add(objListRec); current = null;
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
                if ((objListRec.Option & MTEFConst.mtfeOPT_NUDGE) == MTEFConst.mtfeOPT_NUDGE)
                {
                    dp++;
                    objListRec.NudgeDX = data[dp];
                    dp++;
                    objListRec.NudgeDY = data[dp];
                    if (objListRec.NudgeDX == 0x80 && objListRec.NudgeDY == 0x80)
                    {
                        dp++;
                        dp++;
                        dp++;
                        dp++;
                    }
                }
                if ((objListRec.Option & MTEFConst.mtfeOPT_CHAR_ENC_CHAR_8) != MTEFConst.mtfeOPT_CHAR_ENC_CHAR_8)
                {
                    // typeface
                    dp++;
                    objListRec.Selector = data[dp];
                    dp++;
                    objListRec.Variation = data[dp];
                    objListRec.VariationStr = UnicodeToString(data, dp, 2);
                    dp++; //skip lower byte
                    Console.WriteLine("Variation = {0:X2}, VariationStr = {1}", objListRec.Variation, objListRec.VariationStr);
                }
                else
                {
                    // typeface
                    dp++;
                    objListRec.Selector = data[dp];
                    dp++;
                    objListRec.VariationStr = UnicodeToString(data, dp, 2);
                    Console.WriteLine("Variation = \\u{0:X2}{1:X2}, VariationStr = {2}", data[dp + 1], data[dp], objListRec.VariationStr);
                    dp++; //Skip MTCode
                    dp++;
                    objListRec.Variation = data[dp];
                }
                if ((objListRec.Option & MTEFConst.mtfeOPT_CHAR_EMBELL) == MTEFConst.mtfeOPT_CHAR_EMBELL)
                {
                    dp++;
                    objListRec.EmbellRec.RecType = data[dp];
                    dp++;
                    objListRec.EmbellRec.Option = data[dp];
                    if ((objListRec.EmbellRec.Option & MTEFConst.mtfeOPT_NUDGE) == MTEFConst.mtfeOPT_NUDGE)
                    {
                        dp++;
                        byte embellNudgeDX = data[dp];
                        dp++;
                        byte embellNudgeDY = data[dp];
                        if (embellNudgeDX == 0x80 && embellNudgeDY == 0x80)
                        {
                            dp++;
                            dp++;
                            dp++;
                            dp++;
                        }
                    }
                    dp++;
                    objListRec.EmbellRec.Embell = data[dp];
                    if (data[dp + 1] == 0x00)
                        dp++;
                    Console.WriteLine("Variation = {0:X2}, VariationStr = {1}, Embell = {2:X2}", objListRec.Variation, objListRec.VariationStr, objListRec.EmbellRec.Embell);
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
                OpenedObjList.Add(objListRec); current = null;
                hasMatrix = true;
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
                dp++;
                objListRec.Rows = data[dp];
                dp++;
                objListRec.Cols = data[dp];
                dp++;
                objListRec.RowParts = data[dp];
                dp++;
                objListRec.ColParts = data[dp];
                if (objListRec.Rows > 0 || objListRec.Cols > 0)
                    while (data[dp + 1] == 0x00)
                    {
                        dp++;
                    }

                if (OpenedObjList.Count > 0)
                {
                    objListRec.ParentRec = (AbstractRecord)OpenedObjList.Last<ObjectListRecord>();
                    OpenedObjList.Last<ObjectListRecord>().ChildRecs.Add(objListRec);
                }
                OpenedObjList.Add(objListRec); current = null;
                hasMatrix = true;
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
                if (data[dp + 1] == 0x00)
                    dp++;
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
                    current = OpenedObjList.Last<ObjectListRecord>();
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
            bool isNot = false;
            bool isSpace = false;
            bool isDSpace = false;
            bool isQSpace = false;
            bool isDQSpace = false;
            isNot |= CompareBytes(data, dataPos, MTEFConst.spc1);
            isNot |= CompareBytes(data, dataPos, MTEFConst.spc2);
            isNot |= CompareBytes(data, dataPos, MTEFConst.spc3);
            isNot |= CompareBytes(data, dataPos, MTEFConst.spc4);
            isNot |= CompareBytes(data, dataPos, MTEFConst.spc5);
            isSpace |= CompareBytes(data, dataPos, MTEFConst.spcS1);
            isDSpace |= CompareBytes(data, dataPos, MTEFConst.spcS2);
            isQSpace |= CompareBytes(data, dataPos, MTEFConst.spcQ1);
            isDQSpace |= CompareBytes(data, dataPos, MTEFConst.spcQ2);
            if (isNot)
                return "";
            else if (isSpace)
                return "~";
            else if (isDSpace)
                return "~~";
            else if (isQSpace)
                return "`";
            else if (isDQSpace)
                return "``";
            else
            {
                if (data[dataPos] == 0x00 && data[dataPos + 1] == 0xF7)
                {
                    if (fontName.Equals("Wingdings 2"))
                    {
                        if (data[dataPos + 2] == 0xA2)
                        {
                            data[dataPos] = 0xA0;
                            data[dataPos + 1] = 0x25;
                        }
                        else
                        {
                            data[dataPos] = 0xA1;
                            data[dataPos + 1] = 0x25;
                        }
                    }
                    else if (fontName.Equals("Wingdings 3"))
                    {
                        if (data[dataPos + 2] == 0x70)
                        {
                            data[dataPos] = 0xB2;
                            data[dataPos + 1] = 0x25;
                        }
                        else
                        {
                            data[dataPos] = 0xB3;
                            data[dataPos + 1] = 0x25;
                        }
                    }
                    else return "";
                }
                return Encoding.Unicode.GetString(data, dataPos, length);
            }
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
