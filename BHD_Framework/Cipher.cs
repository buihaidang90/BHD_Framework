using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;

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


    /*
    Examples
    1. Encryption on SQL Server / decryption in C# code.

    SQL
    DECLARE @passphrase varchar(max) = 'password1234'
    SELECT EncryptByPassPhrase(@passphrase, 'Hello World.')
    Result: 0x010000003296649D6782CFD72B8145A07F2C7D7FE3D8B80CF48DA419E94FABC90EEB928D

    C#
    var passphrase = "password1234";
    var decryptedText = SQLServerCryptoMethod.DecryptByPassPhrase(@passphrase, "0x010000003296649D6782CFD72B8145A07F2C7D7FE3D8B80CF48DA419E94FABC90EEB928D");
    System.Console.WriteLine(decryptedText);
    Result: Hello World.

    2. Encryption in C# code / decryption on SQL Server.

    C#
    var passphrase = "password1234";
    var encrypted = SQLServerCryptoMethod.EncryptByPassPhrase(@passphrase, "Hello World.");
    System.Console.WriteLine(encrypted);
    Result: 0x01000000d743db6ccd7e0e63091fa787c65dead5ea14c440da9ee0f6f60e74520a35c076

    SQL
    DECLARE @passphrase varchar(max) = 'password1234'
    SELECT cast(DecryptByPassPhrase(@passphrase, 0x01000000d743db6ccd7e0e63091fa787c65dead5ea14c440da9ee0f6f60e74520a35c076) as varchar)
    Result: Hello World.
     */
    public sealed class HexString
    {
        private readonly byte[] _byteArray;

        private readonly Regex RegExValidation = new Regex("^[0-9a-fA-F]+$", RegexOptions.Compiled);
        private const string PREFIX = "0x";

        public HexString(string hexString)
        {
            if (string.IsNullOrEmpty(hexString))
                throw new ArgumentException("Input string is null or empty.");

            hexString = RemovePrefix(hexString);

            if (hexString.Length % 2 != 0)
                throw new ArgumentException("Invalid number of hexcharacters.");

            if (!RegExValidation.IsMatch(hexString))
                throw new ArgumentException("Input string does not contain hexadecimal characters.");

            _byteArray = HexStringToByteArray(hexString);
        }

        public HexString(byte[] byteArray)
        {
            if (byteArray == null)
                throw new ArgumentException("Input array is null.");

            if (byteArray.Length == 0)
                throw new ArgumentException("Input array is empty.");

            _byteArray = byteArray;
        }

        public string ValueWithoutPrefix => ByteArrayToHexString(_byteArray);

        public string ValueWithPrefix => PREFIX + ValueWithoutPrefix;

        public override string ToString() => ValueWithPrefix;

        public byte[] ToByteArray() => _byteArray;

        public static implicit operator string(HexString hexString) => hexString.ToString();

        public static implicit operator byte[] (HexString hexString) => hexString._byteArray;

        private static string RemovePrefix(string input) => input.StartsWith(PREFIX) ? input.Remove(0, 2) : input;

        // https://stackoverflow.com/questions/321370/how-can-i-convert-a-hex-string-to-a-byte-array#321404
        // Looks nice but could be faster.
        private static byte[] HexStringToByteArray(string hexString) => Enumerable.Range(0, hexString.Length)
               .Where(x => x % 2 == 0)
               .Select(x => Convert.ToByte(hexString.Substring(x, 2), 16))
               .ToArray();

        private static string ByteArrayToHexString(byte[] byteArray) =>
            BitConverter.ToString(byteArray).Replace("-", string.Empty).ToLower();
    }
    public class SQLServerCryptoAlgorithm
    {
        public readonly SQLServerCryptoVersion Version;

        public readonly HashAlgorithm Hash;

        public readonly SymmetricAlgorithm Symmetric;

        public readonly int KeySize;

        public SQLServerCryptoAlgorithm(SQLServerCryptoVersion sqlCryptoVersion)
        {
            Version = sqlCryptoVersion;
            switch (Version)
            {
                case SQLServerCryptoVersion.V1:
                    Hash = SHA1.Create();
                    Symmetric = TripleDES.Create();
                    KeySize = 16;
                    break;
                case SQLServerCryptoVersion.V2:
                    Hash = SHA256.Create();
                    Symmetric = Aes.Create();
                    KeySize = 32;
                    break;
                default:
                    throw new Exception("Unsupported SQLServerCryptoVersion");
            }
            Symmetric.Padding = PaddingMode.PKCS7;
            Symmetric.Mode = CipherMode.CBC;
        }

        public void SetKeyFromPassPhrase(string passphrase)
            => Symmetric.Key = Hash
                .ComputeHash(Encoding.Unicode.GetBytes(passphrase))
                .Take(KeySize)
                .ToArray();
    }
    public class SQLServerCryptoHeader
    {
        public SQLServerCryptoVersion Version = SQLServerCryptoVersion.V1;
        public byte[] Reserved = { 0, 0, 0 };
        public byte[] InitializationVector = { };

        public SQLServerCryptoHeader(SQLServerCryptoVersion sqlServerCryptoVersion = SQLServerCryptoVersion.V1)
        {
            Version = sqlServerCryptoVersion;
        }

        public static implicit operator byte[] (SQLServerCryptoHeader sqlServerCryptoHeader) => sqlServerCryptoHeader.ToByteArray();

        public byte[] ToByteArray()
        {
            var result = new List<byte>();
            result.Add((byte)Version);
            result.AddRange(Reserved);
            result.AddRange(InitializationVector);
            return result.ToArray();
        }
    }
    public class SQLServerCryptoMessage
    {
        private const uint MAGIC_NUMBER = 0xbaadf00d; //0xbaad_f00d

        public uint MagicNumber { get; private set; }

        public ushort IntegrityBytesLength { get; private set; }

        public ushort PlainTextLength { get; private set; }

        public byte[] IntegrityBytes;

        public byte[] MessageBytes;

        public bool AddAuthenticator = false;

        private string _authenticator;

        public string Authenticator
        {
            get { return _authenticator; }

            set
            {
                _authenticator = value.Length > 128 ? _authenticator.Substring(0, 128) : value;
            }
        }


        public SQLServerCryptoMessage() { MagicNumber = MAGIC_NUMBER; }

        public void CreateFromClearText(string cleartext)
        {
            MessageBytes = Encoding.ASCII.GetBytes(cleartext);

            if (MessageBytes.Length > 8000)
                throw new ArgumentOutOfRangeException("The size of the cleartext string should not exceed 8000 bytes.");

            MagicNumber = MAGIC_NUMBER;
            IntegrityBytesLength = 0;

            PlainTextLength = (ushort)MessageBytes.Length;

            if (AddAuthenticator)
            {
                var integrityMessage = MessageBytes.Concat(Encoding.ASCII.GetBytes(Authenticator)).ToArray();
                IntegrityBytes = SHA1.Create().ComputeHash(integrityMessage);
                IntegrityBytesLength = (ushort)IntegrityBytes.Length;
            }
        }

        public void CreateFromDecryptedMessage(byte[] decryptedMessage, bool verify = true)
        {
            MagicNumber = BitConverter.ToUInt32(decryptedMessage, 0);
            IntegrityBytesLength = BitConverter.ToUInt16(decryptedMessage, 4);
            PlainTextLength = BitConverter.ToUInt16(decryptedMessage, 6);

            var messageWithoutHeader = decryptedMessage.Skip(8);

            if (IntegrityBytesLength > 0 || IntegrityBytesLength < 0xffff)
                IntegrityBytes = messageWithoutHeader.Take(IntegrityBytesLength).ToArray();

            if (IntegrityBytesLength != 0xffff)
                MessageBytes = messageWithoutHeader.Skip(IntegrityBytesLength).ToArray();
            else
                MessageBytes = messageWithoutHeader.ToArray();

            if (verify)
                VerifyMessage();
        }

        private void VerifyMessage()
        {
            if (MagicNumber != MAGIC_NUMBER)
                throw new Exception("Message integrity error. Magic numbers are different.");

            var integrityMessage = MessageBytes.Concat(Encoding.ASCII.GetBytes(Authenticator)).ToArray();
            var hash = SHA1.Create().ComputeHash(integrityMessage);

            if (IntegrityBytes.Length > 0 && !hash.SequenceEqual(IntegrityBytes))
                throw new Exception("Message integrity error. Invalid authenticator.");

            if (PlainTextLength != MessageBytes.Length)
                throw new Exception("Message integrity error. Invalid message length.");
        }

        public static implicit operator byte[] (SQLServerCryptoMessage sqlServerCryptoMessage) => sqlServerCryptoMessage.ToByteArray();

        public byte[] ToByteArray()
        {
            byte[] result;
            using (var memoryStream = new MemoryStream())
            {
                using (var binaryWriter = new BinaryWriter(memoryStream))
                {
                    binaryWriter.Write(MagicNumber);
                    binaryWriter.Write(IntegrityBytesLength);
                    binaryWriter.Write(PlainTextLength);

                    if (IntegrityBytes != null)
                        binaryWriter.Write(IntegrityBytes);

                    if (MessageBytes != null)
                        binaryWriter.Write(MessageBytes);
                }
                result = memoryStream.ToArray();
            }
            return result;
        }
    }
    public enum SQLServerCryptoVersion
    {
        ///<summary>
        /// TripleDES/SHA1
        /// </summary> 
        V1 = 0x01,
        /// <summary>
        /// AES256/SHA256
        /// </summary>
        V2 = 0x02
    }
    public static class SQLServerCryptoMethod
    {
        // SQL Server: https://docs.microsoft.com/en-us/sql/t-sql/functions/encryptbypassphrase-transact-sql
        public static HexString EncryptByPassPhrase(string passphrase, string cleartext, int add_authenticator, string authenticator)
            => EncryptByPassPhrase(passphrase, cleartext, add_authenticator, authenticator, SQLServerCryptoVersion.V1);

        public static HexString EncryptByPassPhrase(string passphrase, string cleartext, SQLServerCryptoVersion sqlServerCryptoVersion)
             => EncryptByPassPhrase(passphrase, cleartext, 0, string.Empty, sqlServerCryptoVersion);

        public static HexString EncryptByPassPhrase(string passphrase, string cleartext)
             => EncryptByPassPhrase(passphrase, cleartext, 0, string.Empty, SQLServerCryptoVersion.V1);

        public static HexString EncryptByPassPhrase(string passphrase, string cleartext, int add_authenticator, string authenticator, SQLServerCryptoVersion sqlServerCryptoVersion)
        {
            var sqlServerCryptoAlgorithm = new SQLServerCryptoAlgorithm(sqlServerCryptoVersion);
            sqlServerCryptoAlgorithm.SetKeyFromPassPhrase(passphrase);

            byte[] header = new SQLServerCryptoHeader()
            {
                Version = sqlServerCryptoVersion,
                InitializationVector = sqlServerCryptoAlgorithm.Symmetric.IV
            };
            var sqlServerCryptoMessage = new SQLServerCryptoMessage()
            {
                AddAuthenticator = add_authenticator > 0,
                Authenticator = authenticator
            };
            sqlServerCryptoMessage.CreateFromClearText(cleartext);

            byte[] message = sqlServerCryptoMessage;

            var encryptedMessage = sqlServerCryptoAlgorithm.Symmetric
                .CreateEncryptor()
                .TransformFinalBlock(message, 0, message.Length);

            return new HexString(header.Concat(encryptedMessage).ToArray());
        }

        // SQL Server: https://docs.microsoft.com/en-us/sql/t-sql/functions/decryptbypassphrase-transact-sql
        public static string DecryptByPassPhrase(string passphrase, string ciphertext)
            => DecryptByPassPhrase(passphrase, new HexString(ciphertext), 0, string.Empty, true);

        public static string DecryptByPassPhrase(string passphrase, string ciphertext, int add_authenticator, string authenticator)
            => DecryptByPassPhrase(passphrase, new HexString(ciphertext), add_authenticator, authenticator, true);

        public static string DecryptByPassPhraseWithoutVerification(string passphrase, string ciphertext)
          => DecryptByPassPhrase(passphrase, new HexString(ciphertext), 0, string.Empty, false);

        public static string DecryptByPassPhrase(string passphrase, HexString ciphertext, int add_authenticator, string authenticator, bool verify)
        {
            byte[] ciphertextBytes = ciphertext.ToByteArray();
            var version = (SQLServerCryptoVersion)ciphertextBytes[0];

            var sqlServerCryptoAlgorithm = new SQLServerCryptoAlgorithm(version);
            sqlServerCryptoAlgorithm.SetKeyFromPassPhrase(passphrase);

            var versionAndReservedSize = 4;
            var ivSize = sqlServerCryptoAlgorithm.KeySize / 2;

            var header = new SQLServerCryptoHeader
            {
                Version = version,
                InitializationVector = ciphertextBytes.Skip(versionAndReservedSize).Take(ivSize).ToArray()
            };
            sqlServerCryptoAlgorithm.Symmetric.IV = header.InitializationVector;

            var encryptedMessage = ciphertextBytes.Skip(versionAndReservedSize + ivSize).ToArray();

            var decryptedMessage = sqlServerCryptoAlgorithm.Symmetric
                .CreateDecryptor()
                .TransformFinalBlock(encryptedMessage, 0, encryptedMessage.Length);

            // Message
            var sqlServerCryptoMessage = new SQLServerCryptoMessage()
            {
                AddAuthenticator = add_authenticator > 0,
                Authenticator = authenticator
            };
            sqlServerCryptoMessage.CreateFromDecryptedMessage(decryptedMessage, verify);

            return ByteArray2String(sqlServerCryptoMessage.MessageBytes);
        }

        private static string ByteArray2String(byte[] array)
            => array.Aggregate(string.Empty, (a, b) => a + Convert.ToChar(b));
    }

}
