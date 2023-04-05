using System;
using System.Collections.Generic;
using System.Security.Cryptography;
using System.Text;

namespace BHD_Framework
{
    public static class Cipher
    {
        #region Encrypt/Decrypt a string with Mankichi algorithm.
        /// <summary>
        /// Encrypt/Decrypt a string by Mankichi algorithm without Password.
        /// </summary>
        /// <param name="inputText">Text which want to encrypt/decrypt</param>
        /// <returns></returns>
        public static string ToggleMankichi(string inputText)
        {
            char x;
            string mystr = "";
            for (int i = 0; i < inputText.Length; i++)
            {
                if (((int)inputText[i] >= 65) && ((int)inputText[i] < 91))
                {
                    x = (char)((int)inputText[i] + 13);
                    if ((int)x >= 91) { x = (char)((int)x - 26); }
                }
                else
                {
                    if (((int)inputText[i] >= 97) && ((int)inputText[i] < 123))
                    {
                        x = (char)((int)inputText[i] + 13);
                        if ((int)x >= 123) { x = (char)((int)x - 26); }
                    }
                    else
                    {
                        x = inputText[i];
                    }
                }
                mystr += Convert.ToString(x);
            }
            return mystr;
        }
        #endregion

        #region Encrypt/Decrypt a string with Password.
        /// <summary>
        /// Encrypt a string with Password.
        /// </summary>
        /// <param name="plainText">String want to be encrypted</param>
        /// <param name="privateKey">Secret Key be used to encryption</param>
        public static string Encrypt(string plainText, string privateKey)
        {
            bool useHashing = true;
            byte[] keyArray;
            byte[] toEncryptArray = UTF8Encoding.UTF8.GetBytes(plainText);

            // Get the key from config file
            //AppSettingsReader settingsReader = new AppSettingsReader();
            //string privateKey = (string)settingsReader.GetValue("SecurityKey", typeof(String));
            //System.Windows.Forms.MessageBox.Show(key);

            //If hashing use get hashcode regards to your key
            if (useHashing)
            {
                MD5CryptoServiceProvider hashmd5 = new MD5CryptoServiceProvider();
                keyArray = hashmd5.ComputeHash(UTF8Encoding.UTF8.GetBytes(privateKey));
                //Always release the resources and flush data
                // of the Cryptographic service provide. Best Practice

                hashmd5.Clear();
            }
            else
                keyArray = UTF8Encoding.UTF8.GetBytes(privateKey);

            TripleDESCryptoServiceProvider tdes = new TripleDESCryptoServiceProvider();
            //set the secret key for the tripleDES algorithm
            tdes.Key = keyArray;
            //mode of operation. there are other 4 modes.
            //We choose ECB(Electronic code Book)
            tdes.Mode = CipherMode.ECB;
            //padding mode(if any extra byte added)

            tdes.Padding = PaddingMode.PKCS7;

            ICryptoTransform cTransform = tdes.CreateEncryptor();
            //transform the specified region of bytes array to resultArray
            byte[] resultArray =
              cTransform.TransformFinalBlock(toEncryptArray, 0,
              toEncryptArray.Length);
            //Release resources held by TripleDes Encryptor
            tdes.Clear();
            //Return the encrypted data into unreadable string format
            return Convert.ToBase64String(resultArray, 0, resultArray.Length);
        }

        /// <summary>
        /// Decrypt a string with Password.
        /// </summary>
        /// <param name="encryptedText">String want to be decrypted</param>
        /// <param name="privateKey">Secret Key be used to decryption</param>
        public static string Decrypt(string encryptedText, string privateKey)
        {
            bool useHashing = true;
            byte[] keyArray;
            //get the byte code of the string
            byte[] toEncryptArray = Convert.FromBase64String(encryptedText);

            // Get your key from config file to open the lock!
            //AppSettingsReader settingsReader = new AppSettingsReader();
            //string privateKey = (string)settingsReader.GetValue("SecurityKey",typeof(String));

            if (useHashing)
            {
                //if hashing was used get the hash code with regards to your key
                MD5CryptoServiceProvider hashmd5 = new MD5CryptoServiceProvider();
                keyArray = hashmd5.ComputeHash(UTF8Encoding.UTF8.GetBytes(privateKey));
                //release any resource held by the MD5CryptoServiceProvider

                hashmd5.Clear();
            }
            else
            {
                //if hashing was not implemented get the byte code of the key
                keyArray = UTF8Encoding.UTF8.GetBytes(privateKey);
            }

            TripleDESCryptoServiceProvider tdes = new TripleDESCryptoServiceProvider();
            //set the secret key for the tripleDES algorithm
            tdes.Key = keyArray;
            //mode of operation. there are other 4 modes. 
            //We choose ECB(Electronic code Book)

            tdes.Mode = CipherMode.ECB;
            //padding mode(if any extra byte added)
            tdes.Padding = PaddingMode.PKCS7;

            ICryptoTransform cTransform = tdes.CreateDecryptor();
            byte[] resultArray = cTransform.TransformFinalBlock(toEncryptArray, 0, toEncryptArray.Length);
            //Release resources held by TripleDes Encryptor                
            tdes.Clear();
            //return the Clear decrypted TEXT
            return UTF8Encoding.UTF8.GetString(resultArray);
        }

        /// <summary>
        /// Encrypt/Decrypt a string by MD5 algorithm.
        /// </summary>
        /// <param name="plainText">Text which want to encrypt</param>
        /// <returns></returns>
        public static string MD5Hash(string plainText)
        {
            StringBuilder hash = new StringBuilder();
            MD5CryptoServiceProvider md5provider = new MD5CryptoServiceProvider();
            byte[] bytes = md5provider.ComputeHash(new UTF8Encoding().GetBytes(plainText));

            for (int i = 0; i < bytes.Length; i++)
            {
                hash.Append(bytes[i].ToString("x2"));
            }
            return hash.ToString();
        }

        /// <summary>
        /// Encrypt/Decrypt a string by SHA256 algorithm.
        /// </summary>
        /// <param name="plainText">Text which want to encrypt</param>
        /// <returns></returns>
        public static string SHA256Hash(string plainText)
        {
            // Create a SHA256   
            using (SHA256 sha256Hash = SHA256.Create())
            {
                // ComputeHash - returns byte array  
                byte[] bytes = sha256Hash.ComputeHash(Encoding.UTF8.GetBytes(plainText));

                // Convert byte array to a string   
                StringBuilder builder = new StringBuilder();
                for (int i = 0; i < bytes.Length; i++)
                {
                    builder.Append(bytes[i].ToString("x2"));
                }
                return builder.ToString();
            }
        }
        #endregion

        #region  Encrypt/Decrypt a string with Password with N times.
        public static string EncryptNTimes(string plainText, string privateKey, int timesNumber)
        {
            if (timesNumber <= 0) timesNumber = 1;
            string result = "";
            for (int i = 0; i < timesNumber; i++)
            {
                result = Encrypt(plainText, privateKey);
                result = ToggleMankichi(result);
            }
            return result;
        }
        public static string DecryptNTimes(string cipherText, string privateKey, int timesNumber)
        {
            if (timesNumber <= 0) timesNumber = 1;
            string result = "";
            for (int i = 0; i < timesNumber; i++)
            {
                result = ToggleMankichi(cipherText);
                result = Decrypt(result, privateKey);
            }
            return result;
        }
        #endregion
    }
}
