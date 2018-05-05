using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Text;

namespace Twad.DotNet.Utility.Json
{
    public class Gzip
    {
        #region 压缩流数据

        /// <summary>
        ///     压缩流数据
        /// </summary>
        /// <param name="sourceStream"></param>
        /// <returns></returns>
        public static byte[] EnCompress(Stream sourceStream)
        {
            var vMemory = new MemoryStream();

            sourceStream.Seek(0, SeekOrigin.Begin);
            vMemory.Seek(0, SeekOrigin.Begin);
            try
            {
                using (var vZipStream = new GZipStream(vMemory, CompressionMode.Compress))
                {
                    var vFileByte = new byte[1024];
                    var vRedLen = 0;
                    do
                    {
                        vRedLen = sourceStream.Read(vFileByte, 0, vFileByte.Length);
                        vZipStream.Write(vFileByte, 0, vRedLen);
                    } while (vRedLen > 0);
                }
            }
            finally
            {
                vMemory.Dispose();
            }
            return vMemory.ToArray();
        }


        /// <summary>
        ///     压缩数据
        /// </summary>
        /// <param name="sourceStreambyte"></param>
        /// <returns></returns>
        public static byte[] EnCompress(byte[] sourceStreambyte)
        {
            using (var vMemory = new MemoryStream(sourceStreambyte))
            {
                return EnCompress(vMemory);
            }
        }

        /// <summary>
        ///     压缩数据
        /// </summary>
        /// <param name="sourceStreamStr"></param>
        /// <returns></returns>
        public static byte[] EnCompress(string sourceStreamStr)
        {
            byte[] sourceStreambyte = { };
            try
            {
                sourceStreambyte = Encoding.Default.GetBytes(sourceStreamStr);
                return EnCompress(sourceStreambyte);
            }
            catch (Exception)
            {
                sourceStreambyte = Encoding.Default.GetBytes("EnCompressError");
                return EnCompress(sourceStreambyte);
            }
        }

        #endregion

        #region 解压数据流

        /// <summary>
        ///     解压数据流
        /// </summary>
        /// <param name="sourceStream"></param>
        /// <returns></returns>
        public static byte[] DeCompressByte(Stream sourceStream)
        {
            byte[] vUnZipByte = null;

            using (var vMemory = new MemoryStream())
            {
                var vUnZipStream = new GZipStream(sourceStream, CompressionMode.Decompress);
                try
                {
                    var vTempByte = new byte[1024];
                    var vRedLen = 0;
                    do
                    {
                        vRedLen = vUnZipStream.Read(vTempByte, 0, vTempByte.Length);
                        vMemory.Write(vTempByte, 0, vRedLen);
                    } while (vRedLen > 0);
                    vUnZipStream.Close();
                }
                finally
                {
                    vUnZipStream.Dispose();
                }
                vUnZipByte = vMemory.ToArray();
            }
            return vUnZipByte;
        }

        /// <summary>
        ///     解压数据流
        /// </summary>
        /// <param name="sourceByte"></param>
        /// <returns></returns>
        public static byte[] DeCompressByte(byte[] sourceByte)
        {
            using (var vMemory = new MemoryStream(sourceByte))
            {
                return DeCompressByte(vMemory);
            }
        }

        /// <summary>
        ///     解压数据流
        /// </summary>
        /// <param name="sourceByte"></param>
        /// <returns></returns>
        public static string DeCompressStr(byte[] sourceByte)
        {
            var rtStr = "DeCompressError";
            try
            {
                var deCompress = DeCompressByte(sourceByte);
                rtStr = Encoding.Default.GetString(deCompress);
                return rtStr;
            }
            catch (Exception)
            {
                return rtStr;
            }
        }

        #endregion
    }
}
