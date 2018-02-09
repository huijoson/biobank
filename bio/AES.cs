using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Security.Cryptography;

namespace BioBank
{
    class AES
    {
        public static string AESencrypt(string text, string key)//加密
        {
            //把string轉成byte
            Byte[] plainTextData = Encoding.Unicode.GetBytes(text);
            Byte[] byte_pwd = Encoding.UTF8.GetBytes(key);
            //把key用MD5加密
            MD5CryptoServiceProvider MD5 = new MD5CryptoServiceProvider();
            Byte[] keyData = MD5.ComputeHash(Encoding.Unicode.GetBytes(key));
            Byte[] IVData = MD5.ComputeHash(Encoding.Unicode.GetBytes("5d1b3k"));

            RijndaelManaged AES = new RijndaelManaged();
            ICryptoTransform transform = AES.CreateEncryptor(keyData, IVData);
            Byte[] outputData = transform.TransformFinalBlock(plainTextData, 0, plainTextData.Length);

            return Convert.ToBase64String(outputData);
        }
        public static string AESdecrypt(string text, string key)//解密
        {
            Byte[] cipherTextData = Convert.FromBase64String(text);
            Byte[] byte_pwd = Encoding.UTF8.GetBytes(key);

            MD5CryptoServiceProvider MD5 = new MD5CryptoServiceProvider();
            Byte[] keyData = MD5.ComputeHash(Encoding.Unicode.GetBytes(key));
            Byte[] IVData = MD5.ComputeHash(Encoding.Unicode.GetBytes("5d1b3k"));

            RijndaelManaged AES = new RijndaelManaged();
            ICryptoTransform transform = AES.CreateDecryptor(keyData, IVData);
            try
            {
                Byte[] outputData = transform.TransformFinalBlock(cipherTextData, 0, cipherTextData.Length);
                return Encoding.Unicode.GetString(outputData);
            }
            catch (Exception ex)
            {
                
                return "0";
            }
        }

        public static string GetMD5(string original)
        {
            MD5CryptoServiceProvider md5 = new MD5CryptoServiceProvider();
            byte[] b = md5.ComputeHash(Encoding.UTF8.GetBytes(original));
            return BitConverter.ToString(b).Replace("-", string.Empty);
        }
    }
}
