using System;
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

        public List<AbstractRecord> Objects
        {
            get
            {
                return objects;
            }
        }

        public String Convert(String path)
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
                    return path;
                }
                else
                {
                    return "This GIF file is not a type of the MathType.";
                }
            }
        }

        public String Convert(FileStream stream)
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
                encDefRec.RecType = MTEFConst.REC_ENCODING_DEF;
                dp++;
                encDefRec.EncName = ByteToString(data, ref dp);
                Console.WriteLine("dp = {0}, EncName = {1}", dp, encDefRec.EncName);
            }
            else if (data[dp] == MTEFConst.REC_FONT_DEF)
            {
                Objects.Add(new FontDefRecord());
                FontDefRecord fontDefRec = (FontDefRecord)Objects.Last<AbstractRecord>();
                fontDefRec.RecType = MTEFConst.REC_FONT_DEF;
                dp++;
                fontDefRec.IndexOfEnc = data[dp];
                dp++;
                fontDefRec.FontName = ByteToString(data, ref dp);
                Console.WriteLine("IndexOfEnd = {0}, FontName = {1}, dataPos = {2}", fontDefRec.IndexOfEnc, fontDefRec.FontName, dp);
            }
            else if (data[dp] == MTEFConst.REC_EQN_PREFS)
            {
                Objects.Add(new EqnPrefsRecord());
                EqnPrefsRecord eqnPrefsRec = (EqnPrefsRecord)Objects.Last<AbstractRecord>();
                eqnPrefsRec.RecType = MTEFConst.REC_EQN_PREFS;
                dp++;
                eqnPrefsRec.Option = data[dp];
                eqnPrefsRec.ParseNibbleData(data, ref dp);
                Console.WriteLine("dp = {0}, data = {1}", dp, data[dp]);
            }
            else return;

            dp++;
            if (dp < data.Length)
            {
                parseMTEFRecords(data, dp);
            }
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
 