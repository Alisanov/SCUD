using System;
using System.Security.Cryptography;
using System.Text;

namespace SCUD
{
    public class Encryption
    {
        public static string GetHash(string plaintext)
        {
            SHA1Managed sha = new SHA1Managed();
            byte[] hash = sha.ComputeHash(Encoding.UTF8.GetBytes(plaintext));
            return Convert.ToBase64String(hash);
        }
    }
}
