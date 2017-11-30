using System;
using System.Linq;
using System.Security;

namespace PicImp
{
    class Encrypting
    {
        static byte[] entropy = System.Text.Encoding.Unicode.GetBytes("What 1$ the secret of your powers?");

        //SecureString securedString = new SecureString();
        //string plaintext = "";
        //plaintext.ToCharArray().ToList().ForEach(securedString.AppendChar);
        //string encryptedString = Encrypting.EncryptString(securedString);

        public static string EncryptString(System.Security.SecureString input)
        {
            byte[] encryptedData = System.Security.Cryptography.ProtectedData.Protect(
                System.Text.Encoding.Unicode.GetBytes(ToInsecureString(input)),
                entropy,
                System.Security.Cryptography.DataProtectionScope.LocalMachine);
            string encrypted = Convert.ToBase64String(encryptedData);
            Array.Clear(encryptedData,0,encryptedData.Length);
            return encrypted;
        }

        public static SecureString DecryptString(string encryptedData)
        {
            SecureString securedString = new SecureString();
            string encryptedString = Encrypting.EncryptString(securedString);
            try
            {
                byte[] decryptedData = System.Security.Cryptography.ProtectedData.Unprotect(
                    Convert.FromBase64String(encryptedData),
                    entropy,
                    System.Security.Cryptography.DataProtectionScope.LocalMachine);
                SecureString ss = ToSecureString(System.Text.Encoding.Unicode.GetString(decryptedData));
                Array.Clear(decryptedData, 0, decryptedData.Length);
                return ss;
            }
            catch
            {
                return new SecureString();
            }
        }

        public static SecureString ToSecureString(string input)
        {
            SecureString secure = new SecureString();
            foreach (char c in input)
            {
                secure.AppendChar(c);
            }
            secure.MakeReadOnly();
            return secure;
        }

        public static string ToInsecureString(SecureString input)
        {
            string returnValue = string.Empty;
            IntPtr ptr = System.Runtime.InteropServices.Marshal.SecureStringToBSTR(input);
            try
            {
                returnValue = System.Runtime.InteropServices.Marshal.PtrToStringBSTR(ptr);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ZeroFreeBSTR(ptr);
            }
            return returnValue;
        }
    }
}
