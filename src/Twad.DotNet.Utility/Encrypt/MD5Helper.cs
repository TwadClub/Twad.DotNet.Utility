using System;
using System.Collections.Generic;
using System.Security.Cryptography;
using System.Text;

namespace Twad.DotNet.Utility.Encrypt
{
    public class MD5Helper
    {
        public static string Encrypt(string text, string salt)
        {
            text += salt;

            return Encrypt(text);
        }

        public static string Encrypt(string text)
        {
            byte[] data = Encoding.UTF8.GetBytes(text);
            var md5 = MD5.Create();
            byte[] cipherData = md5.ComputeHash(data);
            return Convert.ToBase64String(cipherData);
        }
    }
}
