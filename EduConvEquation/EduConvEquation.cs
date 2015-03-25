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
        public String Convert(String path)
        {
            if (!File.Exists(path))
                return "File does not exist.";
            else
            {
                FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read);
                fs.Seek(0, SeekOrigin.Begin);
                fs.ReadByte();
                String stringToLookFor = "ABC";
                byte[] bufferToLookFor = { 0x21, 0xFF, 0x0B };

                int matchCounter = 1; // count matches for nicer output

                byte[] readBuffer = new byte[16384]; // our input buffer
                int bytesRead = 0; // number of bytes read
                int offset = 0; // offset inside read-buffer
                long filePos = 0; // position inside the file before read operation
                while( (bytesRead = fs.Read(readBuffer, offset, readBuffer.Length-offset)) > 0 )
                {
                    for( int i=0; i<bytesRead+offset-bufferToLookFor.Length; i++ )
                    {
                        bool match = true;
                        for (int j=0; j<bufferToLookFor.Length; j++)
                            if (bufferToLookFor[j] != readBuffer[i+j])
                            {
                                match = false;
                                break;
                            }
                        if( match )
                        {
                            return path;
                        }
                    }
                    // store file position before next read
                    filePos = fs.Position;

                    // store the last few characters to ensure matches on "chunk boundaries"
                    offset = bufferToLookFor.Length;
                    for (int i=0; i<offset; i++)
                        readBuffer[i] = readBuffer[readBuffer.Length-offset+i];
                }   
                Console.WriteLine("No match found");
                return "No match found";
            }
        }

        public String Convert(FileStream stream)
        {
            return "ABC";
        }
    }
}
