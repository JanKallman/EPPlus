/*******************************************************************************
 * You may amend and distribute as you like, but don't remove this header!
 *
 * EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
 * See http://www.codeplex.com/EPPlus for details.
 *
 * Copyright (C) 2011  Jan Källman
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.

 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  
 * See the GNU Lesser General Public License for more details.
 *
 * The GNU Lesser General Public License can be viewed at http://www.opensource.org/licenses/lgpl-license.php
 * If you unfamiliar with this license or have questions about it, here is an http://www.gnu.org/licenses/gpl-faq.html
 *
 * All code and executables are provided "as is" with no warranty either express or implied. 
 * The author accepts no liability for any damage or loss of business that this product may cause.
 **************************************************************************************
 * This class is created with the help of the MS-OFFCRYPTO PDF documentation... http://msdn.microsoft.com/en-us/library/cc313071(office.12).aspx
 * Decryption library for Office Open XML files(Lyquidity) and Sminks very nice example 
 * on "Reading compound documents in c#" on Stackoverflow. Many thanks!
 ***************************************************************************************
 * Code change notes:
 * 
 * Author							Change						Date
 *******************************************************************************
 * Jan Källman		Added		10-AUG-2010
 * Jan Källman		License changed GPL-->LGPL 2011-12-16
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices; 
using comTypes=System.Runtime.InteropServices.ComTypes;
using System.IO;
using System.Security.Cryptography;
using System.Xml;
namespace OfficeOpenXml.Utils
{
    internal abstract class EncryptionInfo
    {
        internal short MajorVersion;
        internal short MinorVersion;
        internal abstract void Read(byte[] data);

        internal static EncryptionInfo ReadBinary(byte[] data)
        {
            var majorVersion = BitConverter.ToInt16(data, 0);
            var minorVersion = BitConverter.ToInt16(data, 2);
            EncryptionInfo ret;
            if (majorVersion == 3)
            {
                ret=new EncryptionInfoBinary();
            }
            else if (majorVersion == 4)
            {
                ret = new EncryptionInfoAgile();
            }
            else
            {            
                throw(new NotSupportedException("Unsupported encryption format"));
            }
            ret.MajorVersion = majorVersion;
            ret.MinorVersion = minorVersion;
            ret.Read(data);
            return ret;
        }
    }
    /// <summary>
    /// 
    /// </summary>
    internal class EncryptionInfoAgile : EncryptionInfo
    {
        internal class EncryptionKeyData : XmlHelper
        {
            public EncryptionKeyData (XmlNamespaceManager nsm, XmlNode topNode) :
                base(nsm, topNode)
	        {

	        }
            internal byte[] SaltValue 
            { 
                get
                {
                    var s=GetXmlNodeString("@saltValue");
                    return Convert.FromBase64String(s);
                }
                set
                {
                    SetXmlNodeString("@saltValue",Convert.ToBase64String(value));
                }
            }
            internal string HashAlgorithm 
            { 
                get
                {
                    return GetXmlNodeString("@hashAlgorithm");
                }
                set
                {
                    SetXmlNodeString("@hashAlgorithm",value);
                }
            }
            internal string ChiptherChaining
            { 
                get
                {
                    return GetXmlNodeString("@cipherChaining");
                }
                set
                {
                    SetXmlNodeString("@cipherChaining",value);
                }
            }
            internal string CipherAlgorithm 
            { 
                get
                {
                    return GetXmlNodeString("@cipherAlgorithm");
                }
                set
                {
                    SetXmlNodeString("@cipherAlgorithm",value);
                }
            }
            internal int HashSize
            { 
                get
                {
                    return GetXmlNodeInt("@hashSize");
                }
                set
                {
                    SetXmlNodeString("@hashSize",value.ToString());
                }
            }
            internal int KeyBits
            { 
                get
                {
                    return GetXmlNodeInt("@keyBits");
                }
                set
                {
                    SetXmlNodeString("@keyBits", value.ToString());
                }
            }
            internal int BlockSize
            { 
                get
                {
                    return GetXmlNodeInt("@blockSize");
                }
                set
                {
                    SetXmlNodeString("@blockSize", value.ToString());
                }
            }
            internal int SaltSize
            { 
                get
                {
                    return GetXmlNodeInt("@saltSize");
                }
                set
                {
                    SetXmlNodeString("@saltSize",value.ToString());
                }
            }
        }
        internal class EncryptionDataIntegrity : XmlHelper
        {
            public EncryptionDataIntegrity (XmlNamespaceManager nsm, XmlNode topNode) :
                base(nsm, topNode)
	        {

	        }
            internal byte[] HmacValue
            { 
                get
                {
                    var s=GetXmlNodeString("d:dataIntegrity/@encryptedHmacValue");
                    return Convert.FromBase64String(s);
                }
                set
                {
                    SetXmlNodeString("dataIntegrity/@encryptedHmacValue", Convert.ToBase64String(value));
                }
            }
            internal byte[] HmacKey
            { 
                get
                {
                    var s=GetXmlNodeString("dataIntegrity/@encryptedHmacKey");
                    return Convert.FromBase64String(s);
                }
                set
                {
                    SetXmlNodeString("dataIntegrity/@encryptedHmacKey", Convert.ToBase64String(value));
                }            
            }
        }
        internal class EncryptionKeyEncryptor : EncryptionKeyData
        {
            public EncryptionKeyEncryptor(XmlNamespaceManager nsm, XmlNode topNode) :
                base(nsm, topNode)
	        {

	        }
            internal byte[] EncryptedKeyValue
            { 
                get
                {
                    var s=GetXmlNodeString("@encryptedKeyValue");
                    return Convert.FromBase64String(s);
                }
                set
                {
                    SetXmlNodeString("@encryptedKeyValue", Convert.ToBase64String(value));
                }
            }
            internal byte[] EncryptedVerifierHash
            { 
                get
                {
                    var s=GetXmlNodeString("@encryptedVerifierHashValue");
                    return Convert.FromBase64String(s);
                }
                set
                {
                    SetXmlNodeString("@encryptedVerifierHashValue", Convert.ToBase64String(value));
                }
            }
            internal byte[] EncryptedVerifierHashInput
            { 
                get
                {
                    var s=GetXmlNodeString("@encryptedVerifierHashInput");
                    return Convert.FromBase64String(s);
                }
                set
                {
                    SetXmlNodeString("@encryptedVerifierHashInput", Convert.ToBase64String(value));
                }
            }
            internal byte[] VerifierHashInput { get; set; }
            internal byte[] VerifierHash { get; set; }
            internal byte[] KeyValue { get; set; }
            internal int SpinCount
            { 
                get
                {
                    return GetXmlNodeInt("@spinCount");
                }
                set
                {
                    SetXmlNodeString("@spinCount",value.ToString());
                }
            }
        }

        /***
         * <?xml version="1.0" encoding="UTF-8" standalone="true"?>
            <encryption xmlns:c="http://schemas.microsoft.com/office/2006/keyEncryptor/certificate" xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password" xmlns="http://schemas.microsoft.com/office/2006/encryption">
         *      <keyData saltValue="XmTB/XBGJSbwd/GTKzQv5A==" hashAlgorithm="SHA512" cipherChaining="ChainingModeCBC" cipherAlgorithm="AES" hashSize="64" keyBits="256" blockSize="16" saltSize="16"/>
         *      <dataIntegrity encryptedHmacValue="WWw3Bb2dbcNPMnl9f1o7rO0u7sclWGKTXqBA6rRzKsP2KzWS5T0LxY9qFoC6QE67t/t+FNNtMDdMtE3D1xvT8w==" encryptedHmacKey="p/dVdlJY5Kj0k3jI1HRjqtk4s0Y4HmDAsc8nqZgfxNS7DopAsS3LU/2p3CYoIRObHsnHTAtbueH08DFCYGZURg=="/>
         *          <keyEncryptors>
         *              <keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
         *                  <p:encryptedKey saltValue="EeBtY0QftyOkLztCl7NF0g==" hashAlgorithm="SHA512" cipherChaining="ChainingModeCBC" cipherAlgorithm="AES" hashSize="64" keyBits="256" blockSize="16" saltSize="16" encryptedKeyValue="Z7AO8vHnnPZEb1VqyZLJ6JFc3Mq3E322XPxWXS21fbU=" encryptedVerifierHashValue="G7BxbKnZanldvtsbu51mP9J3f9Wr5vCfCpvWSh5eIJff7Sr3J2DzH1/9aKj9uIpqFQIsLohpRk+oBYDcX7hRgw==" encryptedVerifierHashInput="851eszl5y5rdU1RnTjEWHw==" spinCount="100000"/>
         *              </keyEncryptor>
         *      </keyEncryptors>
         *      </encryption
         * ***/
        internal EncryptionDataIntegrity DataIntegrity { get; set; }
        internal EncryptionKeyData KeyData { get; set; }
        internal List<EncryptionKeyEncryptor> KeyEncryptors
        {
            get;
            private set;
        }

        string _xmlEncryptionDescriptor;
        internal override void Read(byte[] data)
        {
            var byXml=new byte[data.Length-8];
            Array.Copy(data,8,byXml,0,data.Length-8);
            _xmlEncryptionDescriptor = Encoding.UTF8.GetString(byXml);
            ReadFromXml();
        }
        private void ReadFromXml()
        {
            var nt=new NameTable();
            var nsm=new XmlNamespaceManager(nt);
            nsm.AddNamespace("d","http://schemas.microsoft.com/office/2006/encryption");
            nsm.AddNamespace("c", "http://schemas.microsoft.com/office/2006/keyEncryptor/certificate");
            nsm.AddNamespace("p","http://schemas.microsoft.com/office/2006/keyEncryptor/password");
            var xml = new XmlDocument();
            XmlHelper.LoadXmlSafe(xml, _xmlEncryptionDescriptor, Encoding.UTF8);
            var node = xml.SelectSingleNode("/d:encryption/d:keyData",nsm);
            KeyData = new EncryptionKeyData(nsm, node);
            node = xml.SelectSingleNode("/d:encryption/d:dataIntegrity", nsm);
            DataIntegrity = new EncryptionDataIntegrity(nsm, node);
            KeyEncryptors=new List<EncryptionKeyEncryptor>();

            var list = xml.SelectNodes("/d:encryption/d:keyEncryptors/d:keyEncryptor/p:encryptedKey", nsm);
            if (list != null)
            {
                foreach (XmlNode n in list)
                {
                    KeyEncryptors.Add(new EncryptionKeyEncryptor(nsm, n));
                }
            }

        }
    }
    /// <summary>
    /// Handles the EncryptionInfo stream
    /// </summary>
    internal class EncryptionInfoBinary : EncryptionInfo
    {

        
        internal Flags Flags;
        internal uint HeaderSize;
        internal EncryptionHeader Header;
        internal EncryptionVerifier Verifier;
        internal override void Read(byte[] data)
        {
            Flags = (Flags)BitConverter.ToInt32(data, 4);
            HeaderSize = (uint)BitConverter.ToInt32(data, 8);

            /**** EncryptionHeader ****/
            Header = new EncryptionHeader();
            Header.Flags = (Flags)BitConverter.ToInt32(data, 12);
            Header.SizeExtra = BitConverter.ToInt32(data, 16);
            Header.AlgID = (AlgorithmID)BitConverter.ToInt32(data, 20);
            Header.AlgIDHash = (AlgorithmHashID)BitConverter.ToInt32(data, 24);
            Header.KeySize = BitConverter.ToInt32(data, 28);
            Header.ProviderType = (ProviderType)BitConverter.ToInt32(data, 32);
            Header.Reserved1 = BitConverter.ToInt32(data, 36);
            Header.Reserved2 = BitConverter.ToInt32(data, 40);

            byte[] text = new byte[(int)HeaderSize - 34];
            Array.Copy(data, 44, text, 0, (int)HeaderSize - 34);
            Header.CSPName = UTF8Encoding.Unicode.GetString(text);

            int pos = (int)HeaderSize + 12;

            /**** EncryptionVerifier ****/
            Verifier = new EncryptionVerifier();
            Verifier.SaltSize = (uint)BitConverter.ToInt32(data, pos);
            Verifier.Salt = new byte[Verifier.SaltSize];

            Array.Copy(data, pos + 4, Verifier.Salt, 0, Verifier.SaltSize);

            Verifier.EncryptedVerifier = new byte[16];
            Array.Copy(data, pos + 20, Verifier.EncryptedVerifier, 0, 16);

            Verifier.VerifierHashSize = (uint)BitConverter.ToInt32(data, pos + 36);
            Verifier.EncryptedVerifierHash = new byte[Verifier.VerifierHashSize];
            Array.Copy(data, pos + 40, Verifier.EncryptedVerifierHash, 0, Verifier.VerifierHashSize);
        }
        internal byte[] WriteBinary()
        {
            MemoryStream ms=new MemoryStream();
            BinaryWriter bw = new BinaryWriter(ms);

            bw.Write(MajorVersion);
            bw.Write(MinorVersion);
            bw.Write((int)Flags);
            byte[] header = Header.WriteBinary();
            bw.Write((uint)header.Length);
            bw.Write(header);
            bw.Write(Verifier.WriteBinary());

            bw.Flush();
            return ms.ToArray();
        }

    }
    /// <summary>
    /// Encryption Header inside the EncryptionInfo stream
    /// </summary>
    internal class EncryptionHeader
    {
        internal Flags Flags;
        internal int SizeExtra;             //MUST be 0x00000000.
        internal AlgorithmID AlgID;         //MUST be 0x0000660E (AES-128), 0x0000660F (AES-192), or 0x00006610 (AES-256).
        internal AlgorithmHashID AlgIDHash; //MUST be 0x00008004 (SHA-1).
        internal int KeySize;               //MUST be 0x00000080 (AES-128), 0x000000C0 (AES-192), or 0x00000100 (AES-256).
        internal ProviderType ProviderType; //SHOULD<10> be 0x00000018 (AES).
        internal int Reserved1;             //Undefined and MUST be ignored.
        internal int Reserved2;             //MUST be 0x00000000 and MUST be ignored.
        internal string CSPName;            //SHOULD<11> be set to either "Microsoft Enhanced RSA and AES Cryptographic Provider" or "Microsoft Enhanced RSA and AES Cryptographic Provider (Prototype)" as a null-terminated Unicode string.
        internal byte[] WriteBinary()
        {
            MemoryStream ms = new MemoryStream();
            BinaryWriter bw = new BinaryWriter(ms);

            bw.Write((int)Flags);
            bw.Write(SizeExtra);
            bw.Write((int)AlgID);
            bw.Write((int)AlgIDHash);
            bw.Write((int)KeySize);
            bw.Write((int)ProviderType);
            bw.Write(Reserved1);
            bw.Write(Reserved2);
            bw.Write(Encoding.Unicode.GetBytes(CSPName));

            bw.Flush();
            return ms.ToArray();
        }
    }
    /// <summary>
    /// Encryption verifier inside the EncryptionInfo stream
    /// </summary>
    internal class EncryptionVerifier
    {
        internal uint SaltSize;              // An unsigned integer that specifies the size of the Salt field. It MUST be 0x00000010.
        internal byte[] Salt;                //(16 bytes): An array of bytes that specifies the salt value used during password hash generation. It MUST NOT be the same data used for the verifier stored encrypted in the EncryptedVerifier field.
        internal byte[] EncryptedVerifier;   //(16 bytes): MUST be the randomly generated Verifier value encrypted using the algorithm chosen by the implementation.
        internal uint VerifierHashSize;      //(4 bytes): An unsigned integer that specifies the number of bytes needed to contain the hash of the data used to generate the EncryptedVerifier field.
        internal byte[] EncryptedVerifierHash; //(variable): An array of bytes that contains the encrypted form of the hash of the randomly generated Verifier value. The length of the array MUST be the size of the encryption block size multiplied by the number of blocks needed to encrypt the hash of the Verifier. If the encryption algorithm is RC4, the length MUST be 20 bytes. If the encryption algorithm is AES, the length MUST be 32 bytes.
        internal byte[] WriteBinary()
        {
            MemoryStream ms = new MemoryStream();
            BinaryWriter bw = new BinaryWriter(ms);

            bw.Write(SaltSize);
            bw.Write(Salt);
            bw.Write(EncryptedVerifier);
            bw.Write(0x14);                 //Sha1 is 20 bytes  (Encrypted is 32)
            bw.Write(EncryptedVerifierHash);

            bw.Flush();
            return ms.ToArray();
        }
    }
    [Flags]
    internal enum Flags
    {
        Reserved1 = 1,   // (1 bit): MUST be set to zero, and MUST be ignored.
        Reserved2 = 2,   // (1 bit): MUST be set to zero, and MUST be ignored.
        fCryptoAPI= 4,   // (1 bit): A flag that specifies whether CryptoAPI RC4 or [ECMA-376] encryption is used. It MUST be set to 1 unless fExternal is 1. If fExternal is set to 1, it MUST be set to zero.        
        fDocProps = 8,   // (1 bit): MUST be set to zero if document properties are encrypted. Otherwise, it MUST be set to 1. Encryption of document properties is specified in section 2.3.5.4.
        fExternal = 16,  // (1 bit): If extensible encryption is used, it MUST be set to 1. Otherwise, it MUST be set to zero. If this field is set to 1, all other fields in this structure MUST be set to zero.
        fAES      = 32   //(1 bit): If the protected content is an [ECMA-376] document, it MUST be set to 1. Otherwise, it MUST be set to zero. If the fAES bit is set to 1, the fCryptoAPI bit MUST also be set to 1
    }
    internal enum AlgorithmID
    {
        Flags   = 0x00000000,   // Determined by Flags
        RC4     = 0x00006801,   // RC4
        AES128  = 0x0000660E,   // 128-bit AES
        AES192  = 0x0000660F,   // 192-bit AES
        AES256  = 0x00006610    // 256-bit AES
    }
    internal enum AlgorithmHashID
    {
        App =  0x00000000,
        SHA1 = 0x00008004,
    }
    internal enum ProviderType
    {
        Flags=0x00000000,//Determined by Flags
        RC4=0x00000001,
        AES=0x00000018,
    }
    /// <summary>
    /// Handels encrypted Excel documents 
    /// </summary>
    internal class EncryptedPackageHandler
    {        
        [DllImport("ole32.dll")]
        private static extern int StgIsStorageFile(
            [MarshalAs(UnmanagedType.LPWStr)] string pwcsName);
        [DllImport("ole32.dll")]
        private static extern int StgIsStorageILockBytes(
            ILockBytes plkbyt);


        [DllImport("ole32.dll")]
        static extern int StgOpenStorage(
            [MarshalAs(UnmanagedType.LPWStr)] string pwcsName,
            IStorage pstgPriority,
            STGM grfMode,
            IntPtr snbExclude,
            uint reserved,
            out IStorage ppstgOpen);

        [DllImport("ole32.dll")]
        static extern int StgOpenStorageOnILockBytes(
            ILockBytes plkbyt,
            IStorage pStgPriority,
            STGM grfMode,
            IntPtr snbEnclude,
            uint reserved,
            out IStorage ppstgOpen);
        [DllImport("ole32.dll")]
        static extern int CreateILockBytesOnHGlobal(
            IntPtr hGlobal,
            bool fDeleteOnRelease,
            out ILockBytes ppLkbyt);

        [DllImport("ole32.dll")]
        static extern int StgCreateDocfileOnILockBytes(ILockBytes plkbyt, STGM grfMode, int reserved, out IStorage ppstgOpen);
        internal static int IsStorageFile(string Name)
        {
            return StgIsStorageFile(Name);
        }
        internal static int IsStorageILockBytes(ILockBytes lb)
        {
            return StgIsStorageILockBytes(lb);
        }
        /// <summary>
        /// Read the package from the OLE document and decrypt it using the supplied password
        /// </summary>
        /// <param name="fi">The file</param>
        /// <param name="encryption"></param>
        /// <returns></returns>
        internal MemoryStream DecryptPackage(FileInfo fi, ExcelEncryption encryption)
        {
            MemoryStream ret = null;
            if (StgIsStorageFile(fi.FullName) == 0)
            {
                IStorage storage = null;
                if (StgOpenStorage(
                    fi.FullName,
                    null,
                    STGM.DIRECT | STGM.READ | STGM.SHARE_EXCLUSIVE,
                    IntPtr.Zero,
                    0,
                    out storage) == 0)
                {
                    ret = GetStreamFromPackage(storage, encryption);                    
                    Marshal.ReleaseComObject(storage);
                }
            }
            else
            {
                throw(new InvalidDataException(string.Format("File {0} is not an encrypted package",fi.FullName)));
            }
            return ret;
        }
        /// <summary>
        /// Read the package from the OLE document and decrypt it using the supplied password
        /// </summary>
        /// <param name="stream">The memory stream. </param>
        /// <param name="encryption">The encryption object from the Package</param>
        /// <returns></returns>
        internal MemoryStream DecryptPackage(MemoryStream stream, ExcelEncryption encryption)
        {
            //Create the lockBytes object.
            ILockBytes lb = GetLockbyte(stream);

            MemoryStream ret = null;

            if (StgIsStorageILockBytes(lb) == 0)
            {
                IStorage storage = null;
                if (StgOpenStorageOnILockBytes(
                        lb,
                        null,
                        STGM.DIRECT | STGM.READ | STGM.SHARE_EXCLUSIVE,
                        IntPtr.Zero,
                        0,
                        out storage) == 0)
                {
                    ret = GetStreamFromPackage(storage, encryption);
                }
                Marshal.ReleaseComObject(storage);
            }
            else
            {
                throw (new InvalidDataException("The stream is not an encrypted package"));
            }
            Marshal.ReleaseComObject(lb);

            return ret;
        }
        internal ILockBytes GetLockbyte(MemoryStream stream)
        {
            ILockBytes lb;
            var iret = CreateILockBytesOnHGlobal(IntPtr.Zero, true, out lb);
            byte[] docArray = stream.GetBuffer();
            IntPtr buffer = Marshal.AllocHGlobal(docArray.Length);
            Marshal.Copy(docArray, 0, buffer, docArray.Length);
            UIntPtr readSize;
            lb.WriteAt(0, buffer, docArray.Length, out readSize);
            Marshal.FreeHGlobal(buffer);
            return lb;
        }
        /// <summary>
        /// Encrypts a package
        /// </summary>
        /// <param name="package">The package as a byte array</param>
        /// <param name="encryption">The encryption info from the workbook</param>
        /// <returns></returns>
        internal MemoryStream EncryptPackage(byte[] package, ExcelEncryption encryption)
        {
            byte[] encryptionKey;
            //Create the Encryption Info. This also returns the Encryptionkey
            var encryptionInfo = CreateEncryptionInfo(encryption.Password, 
                    encryption.Algorithm == EncryptionAlgorithm.AES128 ? 
                        AlgorithmID.AES128 : 
                    encryption.Algorithm == EncryptionAlgorithm.AES192 ? 
                        AlgorithmID.AES192 : 
                        AlgorithmID.AES256, out encryptionKey);

            ILockBytes lb;
            var iret = CreateILockBytesOnHGlobal(IntPtr.Zero, true, out lb);

            IStorage storage = null;
            MemoryStream ret = null;

            //Create the document in-memory
            if (StgCreateDocfileOnILockBytes(lb,
                    STGM.CREATE | STGM.READWRITE | STGM.SHARE_EXCLUSIVE | STGM.TRANSACTED, 
                    0,
                    out storage)==0)
            {
                //First create the dataspace storage
                CreateDataSpaces(storage);

                //Create the Encryption info Stream
                comTypes.IStream stream;
                storage.CreateStream("EncryptionInfo", (uint)(STGM.CREATE | STGM.WRITE | STGM.DIRECT | STGM.SHARE_EXCLUSIVE), (uint)0, (uint)0, out stream);
                byte[] ei=encryptionInfo.WriteBinary();
                stream.Write(ei, ei.Length, IntPtr.Zero);
                stream = null;

                //Encrypt the package
                byte[] encryptedPackage=EncryptData(encryptionKey, package, false);

                storage.CreateStream("EncryptedPackage", (uint)(STGM.CREATE | STGM.WRITE | STGM.DIRECT | STGM.SHARE_EXCLUSIVE), (uint)0, (uint)0, out stream);

                //Write Size here
                MemoryStream ms = new MemoryStream();
                BinaryWriter bw = new BinaryWriter(ms);
                bw.Write((ulong)package.LongLength);
                bw.Flush();
                byte[] length = ms.ToArray();
                //Write Encrypted data length first as an unsigned long
                stream.Write(length, length.Length, IntPtr.Zero);
                //And now the Encrypted data
                stream.Write(encryptedPackage, encryptedPackage.Length, IntPtr.Zero);
                stream = null;
                storage.Commit(0);
                lb.Flush();

                //Now copy the unmanaged stream to a byte array --> memory stream
                var statstg = new comTypes.STATSTG();
                lb.Stat(out statstg, 0);
                int size = (int)statstg.cbSize;
                IntPtr buffer = Marshal.AllocHGlobal(size);
                UIntPtr readSize;
                byte[] pack=new byte[size];
                lb.ReadAt(0, buffer, size, out readSize);
                Marshal.Copy(buffer, pack, 0, size);
                Marshal.FreeHGlobal(buffer);

                ret = new MemoryStream();
                ret.Write(pack, 0, size);
            }
            Marshal.ReleaseComObject(storage);
            Marshal.ReleaseComObject(lb);
            return ret;
        }
        #region "Dataspaces Stream methods"
        private void CreateDataSpaces(IStorage storage)
        {
            IStorage dataSpaces;
            storage.CreateStorage("\x06" + "DataSpaces", (uint)(STGM.CREATE | STGM.WRITE | STGM.DIRECT | STGM.SHARE_EXCLUSIVE), 0, 0, out dataSpaces);
            storage.Commit(0);

            //Version Stream
            comTypes.IStream versionStream;
            dataSpaces.CreateStream("Version", (uint)(STGM.CREATE | STGM.WRITE | STGM.DIRECT | STGM.SHARE_EXCLUSIVE), 0, 0, out versionStream);
            byte[] version = CreateVersionStream();
            versionStream.Write(version,version.Length, IntPtr.Zero);

            //DataSpaceMap
            comTypes.IStream dataSpaceMapStream;
            dataSpaces.CreateStream("DataSpaceMap", (uint)(STGM.CREATE | STGM.WRITE | STGM.DIRECT | STGM.SHARE_EXCLUSIVE), 0, 0, out dataSpaceMapStream);
            byte[] dataSpaceMap = CreateDataSpaceMap();
            dataSpaceMapStream.Write(dataSpaceMap, dataSpaceMap.Length, IntPtr.Zero);

            //DataSpaceInfo
            IStorage dataSpaceInfo;
            dataSpaces.CreateStorage("DataSpaceInfo", (uint)(STGM.CREATE | STGM.WRITE | STGM.DIRECT | STGM.SHARE_EXCLUSIVE), 0, 0, out dataSpaceInfo);

            comTypes.IStream strongEncryptionDataSpaceStream;
            dataSpaceInfo.CreateStream("StrongEncryptionDataSpace", (uint)(STGM.CREATE | STGM.WRITE | STGM.DIRECT | STGM.SHARE_EXCLUSIVE), 0, 0, out strongEncryptionDataSpaceStream);
            byte[] strongEncryptionDataSpace = CreateStrongEncryptionDataSpaceStream();
            strongEncryptionDataSpaceStream.Write(strongEncryptionDataSpace, strongEncryptionDataSpace.Length, IntPtr.Zero);
            dataSpaceInfo.Commit(0);

            //TransformInfo
            IStorage tranformInfo;
            dataSpaces.CreateStorage("TransformInfo", (uint)(STGM.CREATE | STGM.WRITE | STGM.DIRECT | STGM.SHARE_EXCLUSIVE), 0, 0, out tranformInfo);

            IStorage strongEncryptionTransform;
            tranformInfo.CreateStorage("StrongEncryptionTransform", (uint)(STGM.CREATE | STGM.WRITE | STGM.DIRECT | STGM.SHARE_EXCLUSIVE), 0, 0, out strongEncryptionTransform);

            comTypes.IStream primaryStream;
            strongEncryptionTransform.CreateStream("\x06Primary", (uint)(STGM.CREATE | STGM.WRITE | STGM.DIRECT | STGM.SHARE_EXCLUSIVE), 0, 0, out primaryStream);
            byte[] primary = CreateTransformInfoPrimary();
            primaryStream.Write(primary, primary.Length, IntPtr.Zero);
            tranformInfo.Commit(0);
            dataSpaces.Commit(0);
        }
        private byte[] CreateStrongEncryptionDataSpaceStream()
        {
            MemoryStream ms = new MemoryStream();
            BinaryWriter bw = new BinaryWriter(ms);

            bw.Write((int)8);       //HeaderLength
            bw.Write((int)1);       //EntryCount

            string tr = "StrongEncryptionTransform";    
            bw.Write((int)tr.Length);
            bw.Write(UTF8Encoding.Unicode.GetBytes(tr + "\0")); // end \0 is for padding
            
            bw.Flush(); 
            return ms.ToArray();
        }
        private byte[] CreateVersionStream()
        {
            MemoryStream ms = new MemoryStream();
            BinaryWriter bw = new BinaryWriter(ms);

            bw.Write((short)0x3C);  //Major
            bw.Write((short)0);     //Minor
            bw.Write(UTF8Encoding.Unicode.GetBytes("Microsoft.Container.DataSpaces"));
            bw.Write((int)1);       //ReaderVersion
            bw.Write((int)1);       //UpdaterVersion
            bw.Write((int)1);       //WriterVersion
        
            bw.Flush();
            return ms.ToArray();
        }
        private byte[] CreateDataSpaceMap()
        {
            MemoryStream ms = new MemoryStream();
            BinaryWriter bw = new BinaryWriter(ms);

            bw.Write((int)8);       //HeaderLength
            bw.Write((int)1);       //EntryCount
            string s1 = "EncryptedPackage";
            string s2 = "StrongEncryptionDataSpace";
            bw.Write((int)s1.Length + s2.Length+0x14);  
            bw.Write((int)1);       //ReferenceComponentCount
            bw.Write((int)0);       //Stream=0
            bw.Write((int)s1.Length*2); //Length s1
            bw.Write(UTF8Encoding.Unicode.GetBytes(s1));
            bw.Write((int)(s2.Length-1) * 2);   //Length s2
            bw.Write(UTF8Encoding.Unicode.GetBytes(s2 + "\0"));   // end \0 is for padding

            bw.Flush();
            return ms.ToArray();
        }
        private byte[] CreateTransformInfoPrimary()
        {
            MemoryStream ms = new MemoryStream();
            BinaryWriter bw = new BinaryWriter(ms);
            string TransformID="{FF9A3F03-56EF-4613-BDD5-5A41C1D07246}";
            string TransformName = "Microsoft.Container.EncryptionTransform";
            bw.Write(TransformID.Length * 2 + 12);
            bw.Write((int)1);
            bw.Write(TransformID.Length * 2);
            bw.Write(UTF8Encoding.Unicode.GetBytes(TransformID));
            bw.Write(TransformName.Length * 2);
            bw.Write(UTF8Encoding.Unicode.GetBytes(TransformName+"\0"));
            bw.Write((int)1);   //ReaderVersion
            bw.Write((int)1);   //UpdaterVersion
            bw.Write((int)1);   //WriterVersion

            bw.Write((int)0);
            bw.Write((int)0);
            bw.Write((int)0);       //CipherMode
            bw.Write((int)4);       //Reserved

            bw.Flush();
            return ms.ToArray();
        }
        #endregion
        /// <summary>
        /// Create an EncryptionInfo object to encrypt a workbook
        /// </summary>
        /// <param name="password">The password</param>
        /// <param name="algID"></param>
        /// <param name="key">The Encryption key</param>
        /// <returns></returns>
        private EncryptionInfoBinary CreateEncryptionInfo(string password, AlgorithmID algID, out byte[] key)
        {
            if (algID == AlgorithmID.Flags || algID == AlgorithmID.RC4)
            {
                throw(new ArgumentException("algID must be AES128, AES192 or AES256"));
            }
            var encryptionInfo = new EncryptionInfoBinary();
            encryptionInfo.MajorVersion = 4;
            encryptionInfo.MinorVersion = 2;
            encryptionInfo.Flags = Flags.fAES | Flags.fCryptoAPI;
            
            //Header
            encryptionInfo.Header = new EncryptionHeader();
            encryptionInfo.Header.AlgID = algID;
            encryptionInfo.Header.AlgIDHash = AlgorithmHashID.SHA1;
            encryptionInfo.Header.Flags = encryptionInfo.Flags;
            encryptionInfo.Header.KeySize = 
                (algID == AlgorithmID.AES128 ? 0x80 : algID == AlgorithmID.AES192 ? 0xC0 : 0x100);
            encryptionInfo.Header.ProviderType = ProviderType.AES;
            encryptionInfo.Header.CSPName = "Microsoft Enhanced RSA and AES Cryptographic Provider\0";
            encryptionInfo.Header.Reserved1 = 0;
            encryptionInfo.Header.Reserved2 = 0;
            encryptionInfo.Header.SizeExtra = 0;
            
            //Verifier
            encryptionInfo.Verifier = new EncryptionVerifier();
            encryptionInfo.Verifier.Salt = new byte[16];

            var rnd = RandomNumberGenerator.Create();
            rnd.GetBytes(encryptionInfo.Verifier.Salt);
            encryptionInfo.Verifier.SaltSize = 0x10;

            key = GetPasswordHashBinary(password, encryptionInfo);
            
            var verifier = new byte[16];
            rnd.GetBytes(verifier);
            encryptionInfo.Verifier.EncryptedVerifier = EncryptData(key, verifier,true);

            //AES = 32 Bits
            encryptionInfo.Verifier.VerifierHashSize = 0x20;
            SHA1 sha= new SHA1Managed();
            var verifierHash = sha.ComputeHash(verifier);

            encryptionInfo.Verifier.EncryptedVerifierHash = EncryptData(key, verifierHash, false);

            return encryptionInfo;
        }
        private byte[] EncryptData(byte[] key, byte[] data, bool useDataSize)
        {
            RijndaelManaged aes = new RijndaelManaged();
            aes.KeySize = key.Length*8;
            aes.Mode = CipherMode.ECB;
            aes.Padding = PaddingMode.Zeros;

            //Encrypt the data
            var crypt = aes.CreateEncryptor(key, null);
            var ms = new MemoryStream();
            var cs = new CryptoStream(ms, crypt, CryptoStreamMode.Write);
            cs.Write(data, 0, data.Length);
            
            cs.FlushFinalBlock();
            
            byte[] ret;
            if (useDataSize)
            {
                ret = new byte[data.Length];
                ms.Seek(0, SeekOrigin.Begin);
                ms.Read(ret, 0, data.Length);  //Truncate any padded Zeros
                return ret;
            }
            else
            {
                return ms.ToArray();
            }
        }

        private MemoryStream GetStreamFromPackage(IStorage storage, ExcelEncryption encryption)
        {
            MemoryStream ret=null;        
            comTypes.STATSTG statstg;
            
            storage.Stat(out statstg, (uint)STATFLAG.STATFLAG_DEFAULT);

            IEnumSTATSTG pIEnumStatStg = null;
            storage.EnumElements(0, IntPtr.Zero, 0, out pIEnumStatStg);

            comTypes.STATSTG[] regelt = { statstg };
            uint fetched = 0;
            uint res = pIEnumStatStg.Next(1, regelt, out fetched);

            //if (regelt[0].pwcsName == "DataSpaces")
            //{
            //    PrintStorage(storage, regelt[0],"");
            //}
            if (res == 0)
            {
                byte[] data;
                EncryptionInfo encryptionInfo = null;
                while (res != 1)
                {
                    switch (statstg.pwcsName)
                    {
                        case "EncryptionInfo":
                            data = GetOleStream(storage, statstg);
                            //File.WriteAllBytes(@"c:\temp\EncInfo1.bin", data);
                            encryptionInfo = EncryptionInfo.ReadBinary(data);
                            //encryptionInfo.ReadBinary(data);

                            break;
                        case "EncryptedPackage":
                            data = GetOleStream(storage, statstg);
                            ret = DecryptDocument(data, encryptionInfo, encryption.Password);
                            break;
                    }

                    if ((res = pIEnumStatStg.Next(1, regelt, out fetched)) != 1)
                    {
                        statstg = regelt[0];
                    }
                }
            }
            Marshal.ReleaseComObject(pIEnumStatStg);
            return ret;
        }
        // Help method to print a storage part binary to c:\temp
        //private void PrintStorage(IStorage storage, System.Runtime.InteropServices.ComTypes.STATSTG sTATSTG, string topName)
        //{
        //    IStorage ds;
        //    if (topName.Length > 0)
        //    {
        //        topName = topName[0] < 'A' ? topName.Substring(1, topName.Length - 1) : topName;
        //    }
        //    storage.OpenStorage(sTATSTG.pwcsName,
        //        null,
        //        (uint)(STGM.DIRECT | STGM.READ | STGM.SHARE_EXCLUSIVE),
        //        IntPtr.Zero,
        //        0,
        //        out ds);

        //    System.Runtime.InteropServices.ComTypes.STATSTG statstgSub;
        //    ds.Stat(out statstgSub, (uint)STATFLAG.STATFLAG_DEFAULT);

        //    IEnumSTATSTG pIEnumStatStgSub = null;
        //    System.Runtime.InteropServices.ComTypes.STATSTG[] regeltSub = { statstgSub };
        //    ds.EnumElements(0, IntPtr.Zero, 0, out pIEnumStatStgSub);

        //    uint fetched = 0;
        //    while (pIEnumStatStgSub.Next(1, regeltSub, out fetched) == 0)
        //    {
        //        string sName = regeltSub[0].pwcsName[0] < 'A' ? regeltSub[0].pwcsName.Substring(1, regeltSub[0].pwcsName.Length - 1) : regeltSub[0].pwcsName;
        //        if (regeltSub[0].type == 1)
        //        {
        //            PrintStorage(ds, regeltSub[0], topName + sName + "_");
        //        }
        //        else if(regeltSub[0].type==2)
        //        {
        //            File.WriteAllBytes(@"c:\temp\" + topName + sName + ".bin", GetOleStream(ds, regeltSub[0]));
        //        }
        //    }
        //}

        /// <summary>
        /// Decrypt a document
        /// </summary>
        /// <param name="data">The Encrypted data</param>
        /// <param name="encryptionInfo">Encryption Info object</param>
        /// <param name="password">The password</param>
        /// <returns></returns>
        private MemoryStream DecryptDocument(byte[] data, EncryptionInfo encryptionInfo, string password)
        {
            if (encryptionInfo == null)
            {
                throw(new InvalidDataException("Invalid document. EncryptionInfo is missing"));
            }
            long size = BitConverter.ToInt64(data,0);

            var encryptedData = new byte[data.Length - 8];
            Array.Copy(data, 8, encryptedData, 0, encryptedData.Length);

            if (encryptionInfo is EncryptionInfoBinary)
            {
                return DecryptBinary((EncryptionInfoBinary)encryptionInfo, password, size, encryptedData);
            }
            else
            {
                return DecryptAgile((EncryptionInfoAgile)encryptionInfo, password, size, encryptedData);
            }
                
        }

        readonly byte[] BlockKey_HashInput = new byte[] { 0xfe, 0xa7, 0xd2, 0x76, 0x3b, 0x4b, 0x9e, 0x79 };
        readonly byte[] BlockKey_HashValue = new byte[] { 0xd7, 0xaa, 0x0f, 0x6d, 0x30, 0x61, 0x34, 0x4e };
        readonly byte[] BlockKey_KeyValue = new byte[] { 0x14, 0x6e, 0x0b, 0xe7, 0xab, 0xac, 0xd0, 0xd6 };

        private MemoryStream DecryptAgile(EncryptionInfoAgile encryptionInfo, string password, long size, byte[] encryptedData)
        {
            MemoryStream doc = new MemoryStream();

            if (encryptionInfo.KeyData.CipherAlgorithm=="AES")
            {
                RijndaelManaged decryptKey = new RijndaelManaged();
                decryptKey.KeySize = encryptionInfo.KeyData.KeyBits;
                decryptKey.Mode = encryptionInfo.KeyData.ChiptherChaining=="CBC" ? CipherMode.CBC : CipherMode.ECB;
                decryptKey.Padding = PaddingMode.None;
                
                var encr = encryptionInfo.KeyEncryptors[0];
                var hashProvider = GetHashProvider(encr);
                var baseHash = GetPasswordHash(hashProvider, encr.SaltValue, password, encr.SpinCount, encr.HashSize);

                //Get the keys for verifiers and the key value
                var valInputKey = GetFinalHash(hashProvider, encr, BlockKey_HashInput, baseHash);
                var valHashKey = GetFinalHash(hashProvider, encr, BlockKey_HashValue, baseHash);
                var valKeySizeKey = GetFinalHash(hashProvider, encr, BlockKey_KeyValue, baseHash);

                //Decrypt
                encr.VerifierHashInput = DecryptAgileFromKey(encr, valInputKey, encr.EncryptedVerifierHashInput, encr.SaltSize, encr.SaltValue);
                encr.VerifierHash = DecryptAgileFromKey(encr, valHashKey, encr.EncryptedVerifierHash, encr.HashSize, encr.SaltValue);
                encr.KeyValue = DecryptAgileFromKey(encr, valKeySizeKey, encr.EncryptedKeyValue, encr.KeyBits / 8, encr.SaltValue);
                
                if(IsPasswordValid(hashProvider, encr))
                {
                    var br = new BinaryWriter(doc);
                    int pos = 0;
                    int segment=0;
                    while(pos < size)
                    {
                        var segmentSize = (int)(size - pos > 4096 ? 4096 : size - pos);
                        var bufferSize = (int)(encryptedData.Length - pos > 4096 ? 4096 : encryptedData.Length - pos);
                        var iv = new byte[encr.BlockSize];
                        Array.Copy(BitConverter.GetBytes(segment),iv,4);
                        Array.Copy(encr.SaltValue,0,iv,4,encr.BlockSize-4);

                        var buffer = new byte[bufferSize];
                        Array.Copy(encryptedData, pos, buffer, 0, segmentSize);
                        
                        var b=DecryptAgileFromKey(encr, encr.KeyValue, buffer, segmentSize, iv);
                        br.Write(b);
                        pos+=segmentSize;
                        segment++;
                    }
                    br.Flush();
                    return doc;
                }
            }
            return null;
        }

        private HashAlgorithm GetHashProvider(EncryptionInfoAgile.EncryptionKeyEncryptor encr)
        {
            HashAlgorithm hashProvider;
            if (encr.HashAlgorithm == "SHA1")
            {
                hashProvider = new SHA1CryptoServiceProvider();
            }
            else if (encr.HashAlgorithm == "SHA256")
            {
                hashProvider = new SHA256CryptoServiceProvider();
            }
            else if (encr.HashAlgorithm == "SHA384")
            {
                hashProvider = new SHA384CryptoServiceProvider();
            }
            else if (encr.HashAlgorithm == "SHA512")
            {
                hashProvider = new SHA512CryptoServiceProvider();
            }
            else
            {
                throw new NotSupportedException("Hash provider is unsupported. Must be SHA1");
            }
            return hashProvider;
        }

        private MemoryStream DecryptBinary(EncryptionInfoBinary encryptionInfo, string password, long size, byte[] encryptedData)
        {
            MemoryStream doc = new MemoryStream();

            if (encryptionInfo.Header.AlgID == AlgorithmID.AES128 || (encryptionInfo.Header.AlgID == AlgorithmID.Flags && ((encryptionInfo.Flags & (Flags.fAES | Flags.fExternal | Flags.fCryptoAPI)) == (Flags.fAES | Flags.fCryptoAPI)))
                ||
                encryptionInfo.Header.AlgID == AlgorithmID.AES192
                ||
                encryptionInfo.Header.AlgID == AlgorithmID.AES256
                )
            {
                RijndaelManaged decryptKey = new RijndaelManaged();
                decryptKey.KeySize = encryptionInfo.Header.KeySize;
                decryptKey.Mode = CipherMode.ECB;
                decryptKey.Padding = PaddingMode.None;

                var key = GetPasswordHashBinary(password, encryptionInfo);

                if (IsPasswordValid(key, encryptionInfo))
                {
                    ICryptoTransform decryptor = decryptKey.CreateDecryptor(
                                                             key,
                                                             null);


                    MemoryStream dataStream = new MemoryStream(encryptedData);

                    CryptoStream cryptoStream = new CryptoStream(dataStream,
                                                                  decryptor,
                                                                  CryptoStreamMode.Read);

                    var decryptedData = new byte[size];

                    cryptoStream.Read(decryptedData, 0, (int)size);
                    doc.Write(decryptedData, 0, (int)size);
                }
                else
                {
                    throw (new UnauthorizedAccessException("Invalid password"));
                }
            }
            return doc;
        }
        /// <summary>
        /// Validate the password
        /// </summary>
        /// <param name="key">The encryption key</param>
        /// <param name="encryptionInfo">The encryption info extracted from the ENCRYPTIOINFO stream inside the OLE document</param>
        /// <returns></returns>
        private bool IsPasswordValid(byte[] key, EncryptionInfoBinary encryptionInfo)
        {
            RijndaelManaged decryptKey = new RijndaelManaged();
            decryptKey.KeySize = encryptionInfo.Header.KeySize;
            decryptKey.Mode = CipherMode.ECB;
            decryptKey.Padding = PaddingMode.None;

            ICryptoTransform decryptor = decryptKey.CreateDecryptor(
                                                     key,
                                                     null);


            //Decrypt the verifier
            MemoryStream dataStream = new MemoryStream(encryptionInfo.Verifier.EncryptedVerifier);
            CryptoStream cryptoStream = new CryptoStream(dataStream,
                                                          decryptor,
                                                          CryptoStreamMode.Read);
            var decryptedVerifier = new byte[16];
            cryptoStream.Read(decryptedVerifier, 0, 16);

            dataStream = new MemoryStream(encryptionInfo.Verifier.EncryptedVerifierHash);

            cryptoStream = new CryptoStream(    dataStream,
                                                decryptor,
                                                CryptoStreamMode.Read);

            //Decrypt the verifier hash
            var decryptedVerifierHash = new byte[16];
            cryptoStream.Read(decryptedVerifierHash, 0, (int)16);

            //Get the hash for the decrypted verifier
            var sha = new SHA1Managed();
            var hash = sha.ComputeHash(decryptedVerifier);

            //Equal?
            for (int i = 0; i < 16; i++)
            {
                if (hash[i] != decryptedVerifierHash[i])
                {
                    return false;
                }
            }
            return true;
        }
        /// <summary>
        /// Validate the password
        /// </summary>
        /// <param name="key">The encryption key</param>
        /// <param name="encryptionInfo">The encryption info extracted from the ENCRYPTIOINFO stream inside the OLE document</param>
        /// <returns></returns>
        private bool IsPasswordValid(HashAlgorithm sha, EncryptionInfoAgile.EncryptionKeyEncryptor encr)
        {
            var valHash = sha.ComputeHash(encr.VerifierHashInput);

            //Equal?
            for (int i = 0; i < valHash.Length; i++)
            {
                if (encr.VerifierHash[i] != valHash[i])
                {
                    return false;
                }
            }
            return true;
        }

        private byte[] DecryptAgileFromKey(EncryptionInfoAgile.EncryptionKeyEncryptor encr, byte[] key, byte[] encryptedData, long size, byte[] iv)
        {
            RijndaelManaged decryptKey = new RijndaelManaged();
            decryptKey.BlockSize = encr.BlockSize << 3;
            decryptKey.KeySize = encr.KeyBits;
            decryptKey.Mode = CipherMode.CBC;
            decryptKey.Padding = PaddingMode.Zeros;
            
            ICryptoTransform decryptor = decryptKey.CreateDecryptor(
                                                        FixHashSize(key,encr.KeyBits/8),
                                                        iv);


            MemoryStream dataStream = new MemoryStream(encryptedData);

            CryptoStream cryptoStream = new CryptoStream(dataStream,
                                                            decryptor,
                                                            CryptoStreamMode.Read);

            var decryptedData = new byte[size];

            cryptoStream.Read(decryptedData, 0, (int)size);
            return decryptedData;
        }

        /// <summary>
        /// Read the stream and return it as a byte-array
        /// </summary>
        /// <param name="storage"></param>
        /// <param name="statstg"></param>
        /// <returns></returns>
        private byte[] GetOleStream(IStorage storage, comTypes.STATSTG statstg)
        {
            comTypes.IStream pIStream;
            storage.OpenStream(statstg.pwcsName,
               IntPtr.Zero,
               (uint)(STGM.READ | STGM.SHARE_EXCLUSIVE),
               0,
               out pIStream);

            byte[] data = new byte[statstg.cbSize];
            pIStream.Read(data, (int)statstg.cbSize, IntPtr.Zero);
            Marshal.ReleaseComObject(pIStream);

            return data;
        } 
        /// <summary>
        /// Create the hash.
        /// This method is written with the help of Lyquidity library, many thanks for this nice sample
        /// </summary>
        /// <param name="password">The password</param>
        /// <param name="encryptionInfo">The encryption info extracted from the ENCRYPTIOINFO stream inside the OLE document</param>
        /// <returns>The hash to encrypt the document</returns>
        private byte[] GetPasswordHashBinary(string password, EncryptionInfoBinary encryptionInfo)
        {
            byte[] hash = null;
            byte[] tempHash = new byte[4+20];    //Iterator + prev. hash
            try
            {
                HashAlgorithm hashProvider;
                if (encryptionInfo.Header.AlgIDHash == AlgorithmHashID.SHA1 || encryptionInfo.Header.AlgIDHash == AlgorithmHashID.App && (encryptionInfo.Flags & Flags.fExternal) == 0)
                {
                    hashProvider = new SHA1CryptoServiceProvider();
                }
                else if (encryptionInfo.Header.KeySize > 0 && encryptionInfo.Header.KeySize < 80)
                {
                    throw new NotSupportedException("RC4 Hash provider is not supported. Must be SHA1(AlgIDHash == 0x8004)");
                }
                else
                {
                    throw new NotSupportedException("Hash provider is invalid. Must be SHA1(AlgIDHash == 0x8004)");
                }

                hash = GetPasswordHash(hashProvider, encryptionInfo.Verifier.Salt, password,50000, 20);

                // Append "block" (0)
                Array.Copy(hash, tempHash, hash.Length);
                Array.Copy(System.BitConverter.GetBytes(0), 0, tempHash, hash.Length, 4);
                hash = hashProvider.ComputeHash(tempHash);

                /***** Now use the derived key algorithm *****/
                byte[] derivedKey = new byte[64];
                int keySizeBytes = encryptionInfo.Header.KeySize / 8;

                //First XOR hash bytes with 0x36 and fill the rest with 0x36
                for (int i = 0; i < derivedKey.Length; i++)
                    derivedKey[i] = (byte)(i < hash.Length ? 0x36 ^ hash[i] : 0x36);


                byte[] X1 = hashProvider.ComputeHash(derivedKey);

                //if verifier size is bigger than the key size we can return X1
                if ((int)encryptionInfo.Verifier.VerifierHashSize > keySizeBytes)
                    return FixHashSize(X1, keySizeBytes);

                //Else XOR hash bytes with 0x5C and fill the rest with 0x5C
                for (int i = 0; i < derivedKey.Length; i++)
                    derivedKey[i] = (byte)(i < hash.Length ? 0x5C ^ hash[i] : 0x5C);

                byte[] X2 = hashProvider.ComputeHash(derivedKey);

                //Join the two and return 
                byte[] join = new byte[X1.Length + X2.Length];

                Array.Copy(X1, 0, join, 0, X1.Length);
                Array.Copy(X2, 0, join, X1.Length, X2.Length);


                return FixHashSize(join, keySizeBytes); 
            }
            catch (Exception ex)
            {
                throw (new Exception("An error occured when the encryptionkey was created", ex));
            }
        }
        /// <summary>
        /// Create the hash.
        /// This method is written with the help of Lyquidity library, many thanks for this nice sample
        /// </summary>
        /// <param name="password">The password</param>
        /// <param name="encryptionInfo">The encryption info extracted from the ENCRYPTIOINFO stream inside the OLE document</param>
        /// <param name="blockKey">The block key appended to the hash to obtain the final hash</param>
        /// <returns>The hash to encrypt the document</returns>
        private byte[] GetPasswordHashAgile(string password, EncryptionInfoAgile.EncryptionKeyEncryptor encr, byte[] blockKey)
        {
            try
            {
                var hashProvider = GetHashProvider(encr);
                var hash=GetPasswordHash(hashProvider, encr.SaltValue, password, encr.SpinCount, encr.HashSize);
                var hashFinal = GetFinalHash(hashProvider, encr, blockKey, hash);

                return FixHashSize(hashFinal, encr.KeyBits/8);
            }
            catch (Exception ex)
            {
                throw (new Exception("An error occured when the encryptionkey was created", ex));
            }
        }

        private byte[] GetFinalHash(HashAlgorithm hashProvider, EncryptionInfoAgile.EncryptionKeyEncryptor encr, byte[] blockKey, byte[] hash)
        {
            //2.3.4.13 MS-OFFCRYPTO
            var tempHash = new byte[encr.HashSize + blockKey.Length];
            Array.Copy(hash, tempHash, encr.HashSize);
            Array.Copy(blockKey, 0, tempHash, encr.HashSize, blockKey.Length);
            var hashFinal = hashProvider.ComputeHash(tempHash);
            return hashFinal;
        }
        private byte[] GetPasswordHash(HashAlgorithm hashProvider, byte[] salt, string password, int spinCount, int hashSize)
        {
            byte[] hash = null;
            byte[] tempHash = new byte[4 + hashSize];    //Iterator + prev. hash
            hash = hashProvider.ComputeHash(CombinePassword(salt, password));

            //Iterate "spinCount" times, inserting i in first 4 bytes and then the prev. hash in byte 5-24
            for (int i = 0; i < spinCount; i++)
            {
                Array.Copy(BitConverter.GetBytes(i), tempHash, 4);
                Array.Copy(hash, 0, tempHash, 4, hash.Length);

                hash = hashProvider.ComputeHash(tempHash);
            }

            return hash;
        }
        private byte[] FixHashSize(byte[] hash, int size)
        {
            byte[] buff = new byte[size];
            Array.Copy(hash, buff, size);
            return buff;
        }
        private byte[] CombinePassword(byte[] salt, string password)
        {
            if (password == "")
            {
                password = "VelvetSweatshop";   //Used if Password is blank
            }
            // Convert password to unicode...
            byte[] passwordBuf = UnicodeEncoding.Unicode.GetBytes(password);
            
            byte[] inputBuf = new byte[salt.Length + passwordBuf.Length];
            Array.Copy(salt, inputBuf, salt.Length);
            Array.Copy(passwordBuf, 0, inputBuf, salt.Length, passwordBuf.Length);            
            return inputBuf;
        }
        internal static ushort CalculatePasswordHash(string Password)
        {
            //Calculate the hash
            //Thanks to Kohei Yoshida for the sample http://kohei.us/2008/01/18/excel-sheet-protection-password-hash/
            ushort hash = 0;
            for (int i = Password.Length - 1; i >= 0; i--)
            {
                hash ^= Password[i];
                hash = (ushort)(((ushort)((hash >> 14) & 0x01))
                                |
                                ((ushort)((hash << 1) & 0x7FFF)));  //Shift 1 to the left. Overflowing bit 15 goes into bit 0
            }

            hash ^= (0x8000 | ('N' << 8) | 'K'); //Xor NK with high bit set(0xCE4B)
            hash ^= (ushort)Password.Length;

            return hash;
        }
    }
    [ComImport] 
    [Guid("0000000d-0000-0000-C000-000000000046")] 
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)] 
    internal interface IEnumSTATSTG 
    { 
        // The user needs to allocate an STATSTG array whose size is celt. 
        [PreserveSig] 
        uint Next( 
            uint celt, 
            [MarshalAs(UnmanagedType.LPArray), Out] 
            System.Runtime.InteropServices.ComTypes.STATSTG[] rgelt, 
            out uint pceltFetched 
        ); 
 
        void Skip(uint celt); 
 
        void Reset(); 
 
        [return: MarshalAs(UnmanagedType.Interface)] 
        IEnumSTATSTG Clone(); 
    } 
 
    [ComImport] 
    [Guid("0000000b-0000-0000-C000-000000000046")] 
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)] 
    interface IStorage 
    { 
        void CreateStream( 
            /* [string][in] */ string pwcsName, 
            /* [in] */ uint grfMode, 
            /* [in] */ uint reserved1, 
            /* [in] */ uint reserved2, 
            /* [out] */ out comTypes.IStream ppstm); 
        
        void OpenStream( 
            /* [string][in] */ string pwcsName, 
            /* [unique][in] */ IntPtr reserved1, 
            /* [in] */ uint grfMode, 
            /* [in] */ uint reserved2,
            /* [out] */ out comTypes.IStream ppstm); 
 
        void CreateStorage( 
            /* [string][in] */ string pwcsName, 
            /* [in] */ uint grfMode, 
            /* [in] */ uint reserved1, 
            /* [in] */ uint reserved2, 
            /* [out] */ out IStorage ppstg); 
 
        void OpenStorage( 
            /* [string][unique][in] */ string pwcsName, 
            /* [unique][in] */ IStorage pstgPriority, 
            /* [in] */ uint grfMode, 
            /* [unique][in] */ IntPtr snbExclude, 
            /* [in] */ uint reserved, 
            /* [out] */ out IStorage ppstg);

        void CopyTo(
            [InAttribute] uint ciidExclude,
            [InAttribute] Guid[] rgiidExclude,
            [InAttribute] IntPtr snbExclude,
            [InAttribute] IStorage pstgDest
        );
 
        void MoveElementTo( 
            /* [string][in] */ string pwcsName, 
            /* [unique][in] */ IStorage pstgDest, 
            /* [string][in] */ string pwcsNewName, 
            /* [in] */ uint grfFlags); 
 
        void Commit( 
            /* [in] */ uint grfCommitFlags); 
 
        void Revert(); 
 
        void EnumElements( 
            /* [in] */ uint reserved1, 
            /* [size_is][unique][in] */ IntPtr reserved2, 
            /* [in] */ uint reserved3, 
            /* [out] */ out IEnumSTATSTG ppenum); 
 
        void DestroyElement( 
            /* [string][in] */ string pwcsName); 
 
        void RenameElement( 
            /* [string][in] */ string pwcsOldName, 
            /* [string][in] */ string pwcsNewName); 
 
        void SetElementTimes( 
            /* [string][unique][in] */ string pwcsName, 
            /* [unique][in] */ System.Runtime.InteropServices.ComTypes.FILETIME pctime, 
            /* [unique][in] */ System.Runtime.InteropServices.ComTypes.FILETIME patime, 
            /* [unique][in] */ System.Runtime.InteropServices.ComTypes.FILETIME pmtime); 
 
        void SetClass( 
            /* [in] */ Guid clsid); 
 
        void SetStateBits( 
            /* [in] */ uint grfStateBits, 
            /* [in] */ uint grfMask); 
 
        void Stat( 
            /* [out] */ out System.Runtime.InteropServices.ComTypes.STATSTG pstatstg, 
            /* [in] */ uint grfStatFlag); 
 
    }
    [ComVisible(false)]
    [ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("0000000A-0000-0000-C000-000000000046")]
    internal interface ILockBytes
    {
        void ReadAt(long ulOffset, System.IntPtr pv, int cb, out UIntPtr pcbRead);
        void WriteAt(long ulOffset, System.IntPtr pv, int cb, out UIntPtr pcbWritten);
        void Flush();
        void SetSize(long cb);
        void LockRegion(long libOffset, long cb, int dwLockType);
        void UnlockRegion(long libOffset, long cb, int dwLockType);
        void Stat(out System.Runtime.InteropServices.ComTypes.STATSTG pstatstg, int grfStatFlag);
    }
    [Flags] 
    internal enum STGM : int 
    { 
        DIRECT = 0x00000000, 
        TRANSACTED = 0x00010000, 
        SIMPLE = 0x08000000, 
        READ = 0x00000000, 
        WRITE = 0x00000001, 
        READWRITE = 0x00000002, 
        SHARE_DENY_NONE = 0x00000040, 
        SHARE_DENY_READ = 0x00000030, 
        SHARE_DENY_WRITE = 0x00000020, 
        SHARE_EXCLUSIVE = 0x00000010, 
        PRIORITY = 0x00040000, 
        DELETEONRELEASE = 0x04000000, 
        NOSCRATCH = 0x00100000, 
        CREATE = 0x00001000, 
        CONVERT = 0x00020000, 
        FAILIFTHERE = 0x00000000, 
        NOSNAPSHOT = 0x00200000, 
        DIRECT_SWMR = 0x00400000, 
    } 
 
    internal enum STATFLAG : uint 
    { 
        STATFLAG_DEFAULT = 0, 
        STATFLAG_NONAME = 1, 
        STATFLAG_NOOPEN = 2 
    } 
 
    internal enum STGTY : int 
    { 
        STGTY_STORAGE = 1, 
        STGTY_STREAM = 2, 
        STGTY_LOCKBYTES = 3, 
        STGTY_PROPERTY = 4 
    }
}
