using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography;
using System.Text;

namespace Twad.DotNet.Utility.Encrypt
{
    public static class SecurityHelper
    {

        /// <summary>
        ///  对称加密算法 
        ///  以ToBase64String 64为字节为基础的数据编码
        /// </summary>
        /// <param name="source">加密源</param>
        /// <param name="key">密钥</param>
        /// <param name="iv">向量</param>
        /// <param name="symmProv">采用哪种加密算法  DES，RC2，Rijndael，TripleDES 有效</param>
        /// <returns>ToBase64String</returns>
        public static string Encrypt(this string source, string key, string iv)
        {
            DESCryptoServiceProvider provider = new DESCryptoServiceProvider();
            byte[] bytes = Encoding.Default.GetBytes(source);
            provider.Key = Encoding.ASCII.GetBytes(key);
            provider.IV = Encoding.ASCII.GetBytes(iv);
            MemoryStream stream = new MemoryStream();
            CryptoStream stream2 = new CryptoStream(stream, provider.CreateEncryptor(), CryptoStreamMode.Write);
            stream2.Write(bytes, 0, bytes.Length);
            stream2.FlushFinalBlock();
            StringBuilder builder = new StringBuilder();
            foreach (byte num in stream.ToArray())
            {
                builder.AppendFormat("{0:X2}", num);
            }
            builder.ToString();
            return builder.ToString();

        }

        /// <summary>
        ///  对称解密算法 只对有效RijndaelManaged、DESCryptoServiceProvider、
        ///  RC2CryptoServiceProvider 和 TripleDESCryptoServiceProvider
        ///  以ToBase64String 64为字节为基础的数据编码
        /// </summary>
        /// <param name="source">加密源</param>
        /// <param name="key">密钥</param>
        /// <param name="iv">向量</param>
        /// <param name="symmProv">采用哪种加密算法  DES，RC2，Rijndael，TripleDES 有效</param>
        /// <returns>ToBase64String</returns>
        public static string Dencrypt(this string source, string key, string iv)
        {
            DESCryptoServiceProvider provider = new DESCryptoServiceProvider();
            byte[] buffer = new byte[source.Length / 2];
            for (int i = 0; i < (source.Length / 2); i++)
            {
                int num2 = Convert.ToInt32(source.Substring(i * 2, 2), 0x10);
                buffer[i] = (byte)num2;
            }
            provider.Key = Encoding.ASCII.GetBytes(key);
            provider.IV = Encoding.ASCII.GetBytes(iv);
            MemoryStream stream = new MemoryStream();
            CryptoStream stream2 = new CryptoStream(stream, provider.CreateDecryptor(), CryptoStreamMode.Write);
            stream2.Write(buffer, 0, buffer.Length);
            stream2.FlushFinalBlock();
            StringBuilder builder = new StringBuilder();
            return Encoding.Default.GetString(stream.ToArray());

        }
    }
}
