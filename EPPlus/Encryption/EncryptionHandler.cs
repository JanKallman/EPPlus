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
 *
 * Code change notes:
 * 
 * Author							Change						Date
 * ******************************************************************************
 * Jan Källman		    Added       		        2013-01-05
 *******************************************************************************/
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Xml;
using comTypes = System.Runtime.InteropServices.ComTypes;

namespace OfficeOpenXml.Encryption
{
    /// <summary>
    /// Handels encrypted Excel documents 
    /// </summary>
    internal class EncryptedPackageHandler
    {
        /// <summary>
        /// Read the package from the OLE document and decrypt it using the supplied password
        /// </summary>
        /// <param name="fi">The file</param>
        /// <param name="encryption"></param>
        /// <returns></returns>
        internal MemoryStream DecryptPackage(FileInfo fi, ExcelEncryption encryption)
        {
            CompoundDocument doc = new CompoundDocument(fi);
            
            WriteDoc(doc.Storage,@"c:\temp\xl\");
            MemoryStream ret = null;
            if (CompoundDocument.IsStorageFile(fi.FullName) == 0)
            {
                ret = GetStreamFromPackage(doc, encryption);
            }
            else
            {
                throw (new InvalidDataException(string.Format("File {0} is not an encrypted package", fi.FullName)));
            }
            return ret;
        }

        private void WriteDoc(CompoundDocument.StoragePart storagePart, string p)
        {
            foreach (var store in storagePart.SubStorage)
            {
                string sdir=p + store.Key.Replace((char)6,'x') + "\\";
                Directory.CreateDirectory(sdir);
                WriteDoc(store.Value, sdir);
            }
            foreach (var str in storagePart.DataStreams)
            {
                File.WriteAllBytes(p + str.Key.Replace((char)6, 'x') + ".bin", str.Value);
            }
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
            CompoundDocument.ILockBytes lb=null;
            try
            {
                lb = CompoundDocument.GetLockbyte(stream);

                if (CompoundDocument.IsStorageILockBytes(lb) == 0)
                {
                    var doc = new CompoundDocument(lb);
                    return GetStreamFromPackage(doc, encryption);
                }
                else
                {
                    Marshal.ReleaseComObject(lb);
                    throw (new InvalidDataException("The stream is not an valid/supported encrypted document."));
                }
            }
            catch (Exception ex)
            {                
                throw (ex);
            }
            finally
            {
                Marshal.ReleaseComObject(lb);
                lb = null;
            }

        }
        //internal ILockBytes GetLockbyte(MemoryStream stream)
        //{
        //    ILockBytes lb;
        //    var iret = CreateILockBytesOnHGlobal(IntPtr.Zero, true, out lb);
        //    byte[] docArray = stream.GetBuffer();
        //    IntPtr buffer = Marshal.AllocHGlobal(docArray.Length);
        //    Marshal.Copy(docArray, 0, buffer, docArray.Length);
        //    UIntPtr readSize;
        //    lb.WriteAt(0, buffer, docArray.Length, out readSize);
        //    Marshal.FreeHGlobal(buffer);
        //    return lb;
        //}
        /// <summary>
        /// Encrypts a package
        /// </summary>
        /// <param name="package">The package as a byte array</param>
        /// <param name="encryption">The encryption info from the workbook</param>
        /// <returns></returns>
        internal MemoryStream EncryptPackage(byte[] package, ExcelEncryption encryption)
        {
            if (encryption.Version == EncryptionVersion.Version3) //Standard encryption
            {
                return EncryptPackageBinary(package, encryption);
            }
            else if (encryption.Version == EncryptionVersion.Version4) //Agile encryption
            {
                return EncryptPackageAgile(package, encryption);
            }
            throw(new ArgumentException("Unsupported encryption version."));
        }

        private MemoryStream EncryptPackageAgile(byte[] package, ExcelEncryption encryption)
        {
            var xml= "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n";
            xml += "<encryption xmlns=\"http://schemas.microsoft.com/office/2006/encryption\" xmlns:p=\"http://schemas.microsoft.com/office/2006/keyEncryptor/password\" xmlns:c=\"http://schemas.microsoft.com/office/2006/keyEncryptor/certificate\">";
            xml += "<keyData saltSize=\"16\" blockSize=\"16\" keyBits=\"256\" hashSize=\"64\" cipherAlgorithm=\"AES\" cipherChaining=\"ChainingModeCBC\" hashAlgorithm=\"SHA512\" saltValue=\"\"/>";
            xml += "<dataIntegrity encryptedHmacKey=\"\" encryptedHmacValue=\"\"/>";
            xml += "<keyEncryptors>";
            xml += "<keyEncryptor uri=\"http://schemas.microsoft.com/office/2006/keyEncryptor/password\">";
            xml += "<p:encryptedKey spinCount=\"100000\" saltSize=\"16\" blockSize=\"16\" keyBits=\"256\" hashSize=\"64\" cipherAlgorithm=\"AES\" cipherChaining=\"ChainingModeCBC\" hashAlgorithm=\"SHA512\" saltValue=\"\" encryptedVerifierHashInput=\"\" encryptedVerifierHashValue=\"\" encryptedKeyValue=\"\" />";
            xml += "</keyEncryptor></keyEncryptors></encryption>";
            
            var encryptionInfo = new EncryptionInfoAgile();
            encryptionInfo.ReadFromXml(xml);
            var encr = encryptionInfo.KeyEncryptors[0];
            var rnd = RandomNumberGenerator.Create();
            var s = new byte[16];

            rnd.GetBytes(s);
            encryptionInfo.KeyData.SaltValue = s;

            rnd.GetBytes(s);
            encr.SaltValue = s;

            encr.KeyValue = new byte[encr.KeyBits / 8];
            rnd.GetBytes(encr.KeyValue);

            //Get the passwork key.
            var hashProvider = GetHashProvider(encryptionInfo.KeyEncryptors[0]);
            var baseHash = GetPasswordHash(hashProvider, encr.SaltValue, encryption.Password, encr.SpinCount, encr.HashSize);
            var hashFinal = GetFinalHash(hashProvider, encr, BlockKey_KeyValue, baseHash);
            hashFinal = FixHashSize(hashFinal, encr.KeyBits / 8);

            var encrData = EncryptDataAgile(package, encryptionInfo, hashProvider);

            /**** Data Integrity ****/
            var saltHMAC=new byte[64];
            rnd.GetBytes(saltHMAC);

            SetHMAC(encryptionInfo,hashProvider,saltHMAC, encrData);

            /**** Verifier ****/
            encr.VerifierHashInput = new byte[16];
            rnd.GetBytes(encr.VerifierHashInput);

            encr.VerifierHash = hashProvider.ComputeHash(encr.VerifierHashInput);

            var VerifierInputKey = GetFinalHash(hashProvider, encr, BlockKey_HashInput, baseHash);
            var VerifierHashKey = GetFinalHash(hashProvider, encr, BlockKey_HashValue, baseHash);
            var KeyValueKey = GetFinalHash(hashProvider, encr, BlockKey_KeyValue, baseHash);

            var ms = new MemoryStream();
            EncryptAgileFromKey(encr, VerifierInputKey, encr.VerifierHashInput, 0, encr.VerifierHashInput.Length, encr.SaltValue, ms);
            encr.EncryptedVerifierHashInput = ms.ToArray();

            ms = new MemoryStream(); 
            EncryptAgileFromKey(encr, VerifierHashKey, encr.VerifierHash, 0, encr.VerifierHash.Length, encr.SaltValue, ms);
            encr.EncryptedVerifierHash = ms.ToArray();

            ms = new MemoryStream();
            EncryptAgileFromKey(encr, KeyValueKey, encr.KeyValue, 0, encr.KeyValue.Length, encr.SaltValue, ms);
            encr.EncryptedKeyValue = ms.ToArray();

            xml = encryptionInfo.Xml.OuterXml;

            var byXml = Encoding.UTF8.GetBytes(xml);
            
            ms = new MemoryStream();
            ms.Write(BitConverter.GetBytes((ushort)4), 0, 2); //Major Version
            ms.Write(BitConverter.GetBytes((ushort)4), 0, 2); //Minor Version
            ms.Write(BitConverter.GetBytes((uint)0x40), 0, 4); //Reserved
            ms.Write(byXml,0,byXml.Length);

            var doc = new CompoundDocument();
            CreateDataSpaces(doc);
            doc.Storage.DataStreams.Add("EncryptionInfo", ms.ToArray());
            
            //Encrypt the package
            //byte[] encryptedPackage = EncryptData(encr.KeyValue, package, false);
            //ms = new MemoryStream();
            //ms.Write(BitConverter.GetBytes((ulong)encrData.LongLength), 0, 8);
            //ms.Write(encrData, 0, encrData.Length);
            doc.Storage.DataStreams.Add("EncryptedPackage", encrData);

            ms = new MemoryStream();
            var e=doc.Save();
            ms.Write(e,0,e.Length);
            return ms;
        }

        private byte[] EncryptDataAgile(byte[] data, EncryptionInfoAgile encryptionInfo, HashAlgorithm hashProvider)
        {
            var ke = encryptionInfo.KeyEncryptors[0];
            RijndaelManaged aes = new RijndaelManaged();
            aes.KeySize = ke.KeyBits;
            aes.Mode = CipherMode.CBC;
            aes.Padding = PaddingMode.Zeros;

            int pos=0;
            int segment=0;

            //Encrypt the data
            var ms = new MemoryStream();
            ms.Write(BitConverter.GetBytes(data.LongLength), 0, 8);
            while (pos < data.Length)
            {
                var segmentSize = (int)(data.Length - pos > 4096 ? 4096 : data.Length - pos);
                
                var ivTmp = new byte[4 + encryptionInfo.KeyData.SaltSize];
                Array.Copy(encryptionInfo.KeyData.SaltValue, 0, ivTmp, 0, encryptionInfo.KeyData.SaltSize);
                Array.Copy(BitConverter.GetBytes(segment), 0, ivTmp, encryptionInfo.KeyData.SaltSize, 4);
                var iv=hashProvider.ComputeHash(ivTmp);

                EncryptAgileFromKey(ke, ke.KeyValue, data, pos, segmentSize, iv, ms);
                pos += segmentSize;
                segment++;
            }
            ms.Flush();
            return ms.ToArray();
        }
        // Set the dataintegrity
        private void SetHMAC(EncryptionInfoAgile ei, HashAlgorithm hashProvider, byte[] salt, byte[] data)
        {
            var iv = GetFinalHash(hashProvider, ei.KeyEncryptors[0], BlockKey_HmacKey, ei.KeyData.SaltValue);
            var ms = new MemoryStream();
            EncryptAgileFromKey(ei.KeyEncryptors[0], ei.KeyEncryptors[0].KeyValue, salt, 0L, salt.LongLength, iv, ms);
            ei.DataIntegrity.EncryptedHmacKey = ms.ToArray();
            
            var h = GetHmacProvider(ei.KeyEncryptors[0], salt);
            var hmacValue = h.ComputeHash(data);

            ms = new MemoryStream();
            iv = GetFinalHash(hashProvider, ei.KeyEncryptors[0], BlockKey_HmacValue, ei.KeyData.SaltValue);
            EncryptAgileFromKey(ei.KeyEncryptors[0], ei.KeyEncryptors[0].KeyValue, hmacValue, 0L, hmacValue.LongLength, iv, ms);
            ei.DataIntegrity.EncryptedHmacValue = ms.ToArray();
        }

        private HMAC GetHmacProvider(EncryptionInfoAgile.EncryptionKeyEncryptor ei, byte[] salt)
        {
            switch (ei.HashAlgorithm.ToUpper())
            {
                case "RIPEMD160":
                    return new HMACRIPEMD160(salt);
                case "MD5":
                    return new HMACMD5(salt);              
                case "SHA1":
                    return new HMACSHA1(salt);
                case "SHA256":
                    return new HMACSHA256(salt);
                case "SHA384":
                    return new HMACSHA384(salt);
                case "SHA512":
                    return new HMACSHA512(salt);
                default:
                    throw(new NotSupportedException(string.Format("Hash method {0} not supported.",ei.HashAlgorithm)));
            }
        }

        private MemoryStream EncryptPackageBinary(byte[] package, ExcelEncryption encryption)
        {
            byte[] encryptionKey;
            //Create the Encryption Info. This also returns the Encryptionkey
            var encryptionInfo = CreateEncryptionInfo(encryption.Password,
                    encryption.Algorithm == EncryptionAlgorithm.AES128 ?
                        AlgorithmID.AES128 :
                    encryption.Algorithm == EncryptionAlgorithm.AES192 ?
                        AlgorithmID.AES192 :
                        AlgorithmID.AES256, out encryptionKey);

            //ILockBytes lb;
            //var iret = CreateILockBytesOnHGlobal(IntPtr.Zero, true, out lb);

            //IStorage storage = null;
            //MemoryStream ret = null;

            var doc = new CompoundDocument();
            //Create the document in-memory
            //if (StgCreateDocfileOnILockBytes(lb,
            //        STGM.CREATE | STGM.READWRITE | STGM.SHARE_EXCLUSIVE | STGM.TRANSACTED,
            //        0,
            //        out storage) == 0)
            //{
                //First create the dataspace storage
                //CreateDataSpaces(storage);
                CreateDataSpaces(doc);

                //Create the Encryption info Stream
                //comTypes.IStream stream;                
                //storage.CreateStream("EncryptionInfo", (uint)(STGM.CREATE | STGM.WRITE | STGM.DIRECT | STGM.SHARE_EXCLUSIVE), (uint)0, (uint)0, out stream);
                //byte[] ei = encryptionInfo.WriteBinary();
                //stream.Write(ei, ei.Length, IntPtr.Zero);
                //stream = null;
                doc.Storage.DataStreams.Add("EncryptionInfo", encryptionInfo.WriteBinary());
                //Encrypt the package
                byte[] encryptedPackage = EncryptData(encryptionKey, package, false);
                MemoryStream ms = new MemoryStream();
                ms.Write(BitConverter.GetBytes((ulong)package.LongLength), 0, 8);
                ms.Write(encryptedPackage, 0, encryptedPackage.Length);
                doc.Storage.DataStreams.Add("EncryptedPackage", ms.ToArray());
                //storage.CreateStream("EncryptedPackage", (uint)(STGM.CREATE | STGM.WRITE | STGM.DIRECT | STGM.SHARE_EXCLUSIVE), (uint)0, (uint)0, out stream);

                //Write Size here
                //MemoryStream ms = new MemoryStream();
                //BinaryWriter bw = new BinaryWriter(ms);
                //bw.Write((ulong)package.LongLength);
                //bw.Flush();
                //byte[] length = ms.ToArray();
                ////Write Encrypted data length first as an unsigned long
                //stream.Write(length, length.Length, IntPtr.Zero);
                ////And now the Encrypted data
                //stream.Write(encryptedPackage, encryptedPackage.Length, IntPtr.Zero);
                //stream = null;
                //storage.Commit(0);
                //lb.Flush();

                //Now copy the unmanaged stream to a byte array --> memory stream
                //var statstg = new comTypes.STATSTG();
                //lb.Stat(out statstg, 0);
                //int size = (int)statstg.cbSize;
                //IntPtr buffer = Marshal.AllocHGlobal(size);
                //UIntPtr readSize;
                //byte[] pack = new byte[size];
                //lb.ReadAt(0, buffer, size, out readSize);
                //Marshal.Copy(buffer, pack, 0, size);
                //Marshal.FreeHGlobal(buffer);
                var ret = new MemoryStream();                
                var buffer = doc.Save();
                ret.Write(buffer, 0, buffer.Length);
            //}
            //Marshal.ReleaseComObject(storage);
            //Marshal.ReleaseComObject(lb);
            return ret;
        }
        #region "Dataspaces Stream methods"
        private void CreateDataSpaces(CompoundDocument doc)
        {
            var ds = new CompoundDocument.StoragePart();
            doc.Storage.SubStorage.Add("\x06" + "DataSpaces", ds);
            //IStorage dataSpaces;
            //storage.CreateStorage("\x06" + "DataSpaces", (uint)(STGM.CREATE | STGM.WRITE | STGM.DIRECT | STGM.SHARE_EXCLUSIVE), 0, 0, out dataSpaces);
            //storage.Commit(0);

            var ver=new CompoundDocument.StoragePart();
            ds.DataStreams.Add("Version", CreateVersionStream());

            ////Version Stream
            //comTypes.IStream versionStream;
            ////dataSpaces.CreateStream("Version", (uint)(STGM.CREATE | STGM.WRITE | STGM.DIRECT | STGM.SHARE_EXCLUSIVE), 0, 0, out versionStream);
            //byte[] version = CreateVersionStream();
            //versionStream.Write(version, version.Length, IntPtr.Zero);

            //DataSpaceMap
            //comTypes.IStream dataSpaceMapStream;
            //dataSpaces.CreateStream("DataSpaceMap", (uint)(STGM.CREATE | STGM.WRITE | STGM.DIRECT | STGM.SHARE_EXCLUSIVE), 0, 0, out dataSpaceMapStream);
            //byte[] dataSpaceMap = CreateDataSpaceMap();
            //dataSpaceMapStream.Write(dataSpaceMap, dataSpaceMap.Length, IntPtr.Zero);

            ds.DataStreams.Add("DataSpaceMap", CreateDataSpaceMap());

            //DataSpaceInfo
            //IStorage dataSpaceInfo;
            //dataSpaces.CreateStorage("DataSpaceInfo", (uint)(STGM.CREATE | STGM.WRITE | STGM.DIRECT | STGM.SHARE_EXCLUSIVE), 0, 0, out dataSpaceInfo);
            var dsInfo=new CompoundDocument.StoragePart();
            ds.SubStorage.Add("DataSpaceInfo", dsInfo);

            //comTypes.IStream strongEncryptionDataSpaceStream;
            //dataSpaceInfo.CreateStream("StrongEncryptionDataSpace", (uint)(STGM.CREATE | STGM.WRITE | STGM.DIRECT | STGM.SHARE_EXCLUSIVE), 0, 0, out strongEncryptionDataSpaceStream);
            //byte[] strongEncryptionDataSpace = CreateStrongEncryptionDataSpaceStream();
            //strongEncryptionDataSpaceStream.Write(strongEncryptionDataSpace, strongEncryptionDataSpace.Length, IntPtr.Zero);
            //dataSpaceInfo.Commit(0);
            dsInfo.DataStreams.Add("StrongEncryptionDataSpace", CreateStrongEncryptionDataSpaceStream());
            //TransformInfo
            //IStorage tranformInfo;
            //dataSpaces.CreateStorage("TransformInfo", (uint)(STGM.CREATE | STGM.WRITE | STGM.DIRECT | STGM.SHARE_EXCLUSIVE), 0, 0, out tranformInfo);
            var transInfo=new CompoundDocument.StoragePart();
            ds.SubStorage.Add("TransformInfo", transInfo);

            //IStorage strongEncryptionTransform;
            //tranformInfo.CreateStorage("StrongEncryptionTransform", (uint)(STGM.CREATE | STGM.WRITE | STGM.DIRECT | STGM.SHARE_EXCLUSIVE), 0, 0, out strongEncryptionTransform);
            var strEncTrans=new CompoundDocument.StoragePart();
            transInfo.SubStorage.Add("StrongEncryptionTransform", strEncTrans);
            strEncTrans.DataStreams.Add("\x06Primary", CreateTransformInfoPrimary());
            //comTypes.IStream primaryStream;
            //strongEncryptionTransform.CreateStream("\x06Primary", (uint)(STGM.CREATE | STGM.WRITE | STGM.DIRECT | STGM.SHARE_EXCLUSIVE), 0, 0, out primaryStream);
            //byte[] primary = CreateTransformInfoPrimary();
            //primaryStream.Write(primary, primary.Length, IntPtr.Zero);
            //tranformInfo.Commit(0);
            //dataSpaces.Commit(0);
        }
        private byte[] CreateStrongEncryptionDataSpaceStream()
        {
            MemoryStream ms = new MemoryStream();
            BinaryWriter bw = new BinaryWriter(ms);

            bw.Write((int)8);       //HeaderLength
            bw.Write((int)1);       //EntryCount

            string tr = "StrongEncryptionTransform";
            bw.Write((int)tr.Length*2);
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
            bw.Write((int)(s1.Length + s2.Length)*2 + 0x16);
            bw.Write((int)1);       //ReferenceComponentCount
            bw.Write((int)0);       //Stream=0
            bw.Write((int)s1.Length * 2); //Length s1
            bw.Write(UTF8Encoding.Unicode.GetBytes(s1));
            bw.Write((int)(s2.Length * 2));   //Length s2
            bw.Write(UTF8Encoding.Unicode.GetBytes(s2 + "\0"));   // end \0 is for padding

            bw.Flush();
            return ms.ToArray();
        }
        private byte[] CreateTransformInfoPrimary()
        {
            MemoryStream ms = new MemoryStream();
            BinaryWriter bw = new BinaryWriter(ms);
            string TransformID = "{FF9A3F03-56EF-4613-BDD5-5A41C1D07246}";
            string TransformName = "Microsoft.Container.EncryptionTransform";
            bw.Write(TransformID.Length * 2 + 12);
            bw.Write((int)1);
            bw.Write(TransformID.Length * 2);
            bw.Write(UTF8Encoding.Unicode.GetBytes(TransformID));
            bw.Write(TransformName.Length * 2);
            bw.Write(UTF8Encoding.Unicode.GetBytes(TransformName + "\0"));
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
                throw (new ArgumentException("algID must be AES128, AES192 or AES256"));
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
            encryptionInfo.Verifier.EncryptedVerifier = EncryptData(key, verifier, true);

            //AES = 32 Bits
            encryptionInfo.Verifier.VerifierHashSize = 0x20;
            SHA1 sha = new SHA1Managed();
            var verifierHash = sha.ComputeHash(verifier);

            encryptionInfo.Verifier.EncryptedVerifierHash = EncryptData(key, verifierHash, false);

            return encryptionInfo;
        }
        private byte[] EncryptData(byte[] key, byte[] data, bool useDataSize)
        {
            RijndaelManaged aes = new RijndaelManaged();
            aes.KeySize = key.Length * 8;
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

        private MemoryStream GetStreamFromPackage(CompoundDocument doc, ExcelEncryption encryption)
        {
            var ret = new MemoryStream();
            if(doc.Storage.DataStreams.ContainsKey("EncryptionInfo") ||
               doc.Storage.DataStreams.ContainsKey("EncryptedPackage"))
            {
                var encryptionInfo = EncryptionInfo.ReadBinary(doc.Storage.DataStreams["EncryptionInfo"]);
                
                return DecryptDocument(doc.Storage.DataStreams["EncryptedPackage"], encryptionInfo, encryption.Password);
            }
            else
            {
                throw (new InvalidDataException("Invalid document. EncryptionInfo or EncryptedPackage stream is missing"));
            }
        }

        /// <summary>
        /// Decrypt a document
        /// </summary>
        /// <param name="data">The Encrypted data</param>
        /// <param name="encryptionInfo">Encryption Info object</param>
        /// <param name="password">The password</param>
        /// <returns></returns>
        private MemoryStream DecryptDocument(byte[] data, EncryptionInfo encryptionInfo, string password)
        {
            long size = BitConverter.ToInt64(data, 0);

            var encryptedData = new byte[data.Length - 8];
            Array.Copy(data, 8, encryptedData, 0, encryptedData.Length);

            if (encryptionInfo is EncryptionInfoBinary)
            {
                return DecryptBinary((EncryptionInfoBinary)encryptionInfo, password, size, encryptedData);
            }
            else
            {
                return DecryptAgile((EncryptionInfoAgile)encryptionInfo, password, size, encryptedData, data);
            }

        }

        readonly byte[] BlockKey_HashInput = new byte[] { 0xfe, 0xa7, 0xd2, 0x76, 0x3b, 0x4b, 0x9e, 0x79 };
        readonly byte[] BlockKey_HashValue = new byte[] { 0xd7, 0xaa, 0x0f, 0x6d, 0x30, 0x61, 0x34, 0x4e };
        readonly byte[] BlockKey_KeyValue = new byte[] { 0x14, 0x6e, 0x0b, 0xe7, 0xab, 0xac, 0xd0, 0xd6 };
        readonly byte[] BlockKey_HmacKey = new byte[] { 0x5f, 0xb2, 0xad, 0x01, 0x0c, 0xb9, 0xe1, 0xf6 };//MSOFFCRYPTO 2.3.4.14 section 3
        readonly byte[] BlockKey_HmacValue = new byte[] { 0xa0, 0x67, 0x7f, 0x02, 0xb2, 0x2c, 0x84, 0x33 };//MSOFFCRYPTO 2.3.4.14 section 5
        
        private MemoryStream DecryptAgile(EncryptionInfoAgile encryptionInfo, string password, long size, byte[] encryptedData, byte[] data)
        { 
            MemoryStream doc = new MemoryStream();

            if (encryptionInfo.KeyData.CipherAlgorithm == "AES")
            {
                var encr = encryptionInfo.KeyEncryptors[0];
                var hashProvider = GetHashProvider(encr);
                var baseHash = GetPasswordHash(hashProvider, encr.SaltValue, password, encr.SpinCount, encr.HashSize);

                //Get the keys for the verifiers and the key value
                var valInputKey = GetFinalHash(hashProvider, encr, BlockKey_HashInput, baseHash);
                var valHashKey = GetFinalHash(hashProvider, encr, BlockKey_HashValue, baseHash);
                var valKeySizeKey = GetFinalHash(hashProvider, encr, BlockKey_KeyValue, baseHash);

                //Decrypt
                encr.VerifierHashInput = DecryptAgileFromKey(encr, valInputKey, encr.EncryptedVerifierHashInput, encr.SaltSize, encr.SaltValue);
                encr.VerifierHash = DecryptAgileFromKey(encr, valHashKey, encr.EncryptedVerifierHash, encr.HashSize, encr.SaltValue);
                encr.KeyValue = DecryptAgileFromKey(encr, valKeySizeKey, encr.EncryptedKeyValue, encr.KeyBits / 8, encr.SaltValue);

                if (IsPasswordValid(hashProvider, encr))
                {
                    var ivhmac = GetFinalHash(hashProvider, encr, BlockKey_HmacKey, encryptionInfo.KeyData.SaltValue);
                    var key = DecryptAgileFromKey(encr, encr.KeyValue, encryptionInfo.DataIntegrity.EncryptedHmacKey, encryptionInfo.KeyData.HashSize, ivhmac);

                    ivhmac = GetFinalHash(hashProvider, encr, BlockKey_HmacValue, encryptionInfo.KeyData.SaltValue);
                    var value = DecryptAgileFromKey(encr, encr.KeyValue, encryptionInfo.DataIntegrity.EncryptedHmacValue, encryptionInfo.KeyData.HashSize, ivhmac);

                    var hmca = new HMACSHA512(key);
                    var v2 = hmca.ComputeHash(data);

                    for(int i=0;i<v2.Length;i++)
                    {
                        if(value[i]!=v2[i])
                        {
                            throw(new Exception("Dataintegrity key missmatch"));
                        }
                    }

                    int pos = 0;
                    uint segment = 0;
                    while (pos < size)
                    {
                        var segmentSize = (int)(size - pos > 4096 ? 4096 : size - pos);
                        var bufferSize = (int)(encryptedData.Length - pos > 4096 ? 4096 : encryptedData.Length - pos);
                        var ivTmp = new byte[4+encryptionInfo.KeyData.SaltSize];
                        Array.Copy(encryptionInfo.KeyData.SaltValue, 0, ivTmp, 0, encryptionInfo.KeyData.SaltSize);
                        Array.Copy(BitConverter.GetBytes(segment), 0, ivTmp, encryptionInfo.KeyData.SaltSize, 4);
                        var iv=hashProvider.ComputeHash(ivTmp);
                        var buffer = new byte[bufferSize];
                        Array.Copy(encryptedData, pos, buffer, 0, bufferSize);

                        var b = DecryptAgileFromKey(encr, encr.KeyValue, buffer, segmentSize, iv);
                        doc.Write(b, 0, b.Length);
                        pos += segmentSize;
                        segment++;
                    }
                    doc.Flush();
                    return doc;
                }
            }
            return null;
        }

        private HashAlgorithm GetHashProvider(EncryptionInfoAgile.EncryptionKeyEncryptor encr)
        {
            HashAlgorithm hashProvider;
            switch (encr.HashAlgorithm.ToUpper())
            {
                case "MD5":
                    {
                        return new MD5CryptoServiceProvider();
                    }
                case "RIPEMD160":
                    {
                        return new RIPEMD160Managed();
                    }
                case "SHA1":
                    {
                        return new SHA1CryptoServiceProvider();
                    }
                case "SHA256":
                    {
                        return  new SHA256CryptoServiceProvider();
                    }
                case "SHA384":
                    {
                        return new SHA384CryptoServiceProvider();
                    }
                case "SHA512":
                    {
                        return new SHA512CryptoServiceProvider();
                    }
                default:
                    {
                        throw new NotSupportedException(string.Format("Hash provider is unsupported. {0}", encr.HashAlgorithm));
                    }
            }
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

            cryptoStream = new CryptoStream(dataStream,
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
                                                        FixHashSize(key, encr.KeyBits / 8),
                                                        FixHashSize(iv, encr.BlockSize, 0x36));


            MemoryStream dataStream = new MemoryStream(encryptedData);

            CryptoStream cryptoStream = new CryptoStream(dataStream,
                                                            decryptor,
                                                            CryptoStreamMode.Read);

            var decryptedData = new byte[size];

            cryptoStream.Read(decryptedData, 0, (int)size);
            return decryptedData;
        }
        private void EncryptAgileFromKey(EncryptionInfoAgile.EncryptionKeyEncryptor encr, byte[] key, byte[] data, long pos, long size, byte[] iv,MemoryStream ms)
        {
            RijndaelManaged encryptKey = new RijndaelManaged();
            encryptKey.BlockSize = encr.BlockSize << 3;
            encryptKey.KeySize = encr.KeyBits;
            encryptKey.Mode = CipherMode.CBC;
            encryptKey.Padding = PaddingMode.Zeros;

            ICryptoTransform encryptor = encryptKey.CreateEncryptor(
                                                        FixHashSize(key, encr.KeyBits / 8),
                                                        FixHashSize(iv, 16, 0x36));


            CryptoStream cryptoStream = new CryptoStream(ms,
                                                         encryptor,
                                                         CryptoStreamMode.Write);
            
            var cryptoSize = size % encr.BlockSize == 0 ? size : (size + (encr.BlockSize - (size % encr.BlockSize)));
            var buffer = new byte[size];
            Array.Copy(data, pos, buffer, 0, size);
            cryptoStream.Write(buffer, 0, (int)size);
            while (size % encr.BlockSize != 0)
            {
                cryptoStream.WriteByte(0);
                size++;
            }
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
            byte[] tempHash = new byte[4 + 20];    //Iterator + prev. hash
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

                hash = GetPasswordHash(hashProvider, encryptionInfo.Verifier.Salt, password, 50000, 20);

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
                var hash = GetPasswordHash(hashProvider, encr.SaltValue, password, encr.SpinCount, encr.HashSize);
                var hashFinal = GetFinalHash(hashProvider, encr, blockKey, hash);

                return FixHashSize(hashFinal, encr.KeyBits / 8);
            }
            catch (Exception ex)
            {
                throw (new Exception("An error occured when the encryptionkey was created", ex));
            }
        }

        private byte[] GetFinalHash(HashAlgorithm hashProvider, EncryptionInfoAgile.EncryptionKeyEncryptor encr, byte[] blockKey, byte[] hash)
        {
            //2.3.4.13 MS-OFFCRYPTO
            var tempHash = new byte[hash.Length + blockKey.Length];
            Array.Copy(hash, tempHash, hash.Length);
            Array.Copy(blockKey, 0, tempHash, hash.Length, blockKey.Length);
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
        private byte[] FixHashSize(byte[] hash, int size, byte fill=0)
        {
            if (hash.Length == size)
                return hash;
            else if (hash.Length < size)
            {
                byte[] buff = new byte[size];
                Array.Copy(hash, buff, hash.Length);
                for (int i = hash.Length; i < size; i++)
                {
                    buff[i] = fill;
                }
                return buff;
            }
            else
            {
                byte[] buff = new byte[size];
                Array.Copy(hash, buff, size);
                return buff;
            }
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

        //[ComImport]
        //[Guid("0000000d-0000-0000-C000-000000000046")]
        //[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        //internal interface IEnumSTATSTG
        //{
        //    // The user needs to allocate an STATSTG array whose size is celt. 
        //    [PreserveSig]
        //    uint Next(
        //        uint celt,
        //        [MarshalAs(UnmanagedType.LPArray), Out] 
        //    System.Runtime.InteropServices.ComTypes.STATSTG[] rgelt,
        //        out uint pceltFetched
        //    );

        //    void Skip(uint celt);

        //    void Reset();

        //    [return: MarshalAs(UnmanagedType.Interface)]
        //    IEnumSTATSTG Clone();
        //}

        //[ComImport]
        //[Guid("0000000b-0000-0000-C000-000000000046")]
        //[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        //interface IStorage
        //{
        //    void CreateStream(
        //        /* [string][in] */ string pwcsName,
        //        /* [in] */ uint grfMode,
        //        /* [in] */ uint reserved1,
        //        /* [in] */ uint reserved2,
        //        /* [out] */ out comTypes.IStream ppstm);

        //    void OpenStream(
        //        /* [string][in] */ string pwcsName,
        //        /* [unique][in] */ IntPtr reserved1,
        //        /* [in] */ uint grfMode,
        //        /* [in] */ uint reserved2,
        //        /* [out] */ out comTypes.IStream ppstm);

        //    void CreateStorage(
        //        /* [string][in] */ string pwcsName,
        //        /* [in] */ uint grfMode,
        //        /* [in] */ uint reserved1,
        //        /* [in] */ uint reserved2,
        //        /* [out] */ out IStorage ppstg);

        //    void OpenStorage(
        //        /* [string][unique][in] */ string pwcsName,
        //        /* [unique][in] */ IStorage pstgPriority,
        //        /* [in] */ uint grfMode,
        //        /* [unique][in] */ IntPtr snbExclude,
        //        /* [in] */ uint reserved,
        //        /* [out] */ out IStorage ppstg);

        //    void CopyTo(
        //        [InAttribute] uint ciidExclude,
        //        [InAttribute] Guid[] rgiidExclude,
        //        [InAttribute] IntPtr snbExclude,
        //        [InAttribute] IStorage pstgDest
        //    );

        //    void MoveElementTo(
        //        /* [string][in] */ string pwcsName,
        //        /* [unique][in] */ IStorage pstgDest,
        //        /* [string][in] */ string pwcsNewName,
        //        /* [in] */ uint grfFlags);

        //    void Commit(
        //        /* [in] */ uint grfCommitFlags);

        //    void Revert();

        //    void EnumElements(
        //        /* [in] */ uint reserved1,
        //        /* [size_is][unique][in] */ IntPtr reserved2,
        //        /* [in] */ uint reserved3,
        //        /* [out] */ out IEnumSTATSTG ppenum);

        //    void DestroyElement(
        //        /* [string][in] */ string pwcsName);

        //    void RenameElement(
        //        /* [string][in] */ string pwcsOldName,
        //        /* [string][in] */ string pwcsNewName);

        //    void SetElementTimes(
        //        /* [string][unique][in] */ string pwcsName,
        //        /* [unique][in] */ System.Runtime.InteropServices.ComTypes.FILETIME pctime,
        //        /* [unique][in] */ System.Runtime.InteropServices.ComTypes.FILETIME patime,
        //        /* [unique][in] */ System.Runtime.InteropServices.ComTypes.FILETIME pmtime);

        //    void SetClass(
        //        /* [in] */ Guid clsid);

        //    void SetStateBits(
        //        /* [in] */ uint grfStateBits,
        //        /* [in] */ uint grfMask);

        //    void Stat(
        //        /* [out] */ out System.Runtime.InteropServices.ComTypes.STATSTG pstatstg,
        //        /* [in] */ uint grfStatFlag);

        //}
        //[ComVisible(false)]
        //[ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("0000000A-0000-0000-C000-000000000046")]
        //internal interface ILockBytes
        //{
        //    void ReadAt(long ulOffset, System.IntPtr pv, int cb, out UIntPtr pcbRead);
        //    void WriteAt(long ulOffset, System.IntPtr pv, int cb, out UIntPtr pcbWritten);
        //    void Flush();
        //    void SetSize(long cb);
        //    void LockRegion(long libOffset, long cb, int dwLockType);
        //    void UnlockRegion(long libOffset, long cb, int dwLockType);
        //    void Stat(out System.Runtime.InteropServices.ComTypes.STATSTG pstatstg, int grfStatFlag);
        //}
        //[Flags]
        //internal enum STGM : int
        //{
        //    DIRECT = 0x00000000,
        //    TRANSACTED = 0x00010000,
        //    SIMPLE = 0x08000000,
        //    READ = 0x00000000,
        //    WRITE = 0x00000001,
        //    READWRITE = 0x00000002,
        //    SHARE_DENY_NONE = 0x00000040,
        //    SHARE_DENY_READ = 0x00000030,
        //    SHARE_DENY_WRITE = 0x00000020,
        //    SHARE_EXCLUSIVE = 0x00000010,
        //    PRIORITY = 0x00040000,
        //    DELETEONRELEASE = 0x04000000,
        //    NOSCRATCH = 0x00100000,
        //    CREATE = 0x00001000,
        //    CONVERT = 0x00020000,
        //    FAILIFTHERE = 0x00000000,
        //    NOSNAPSHOT = 0x00200000,
        //    DIRECT_SWMR = 0x00400000,
        //}

        //internal enum STATFLAG : uint
        //{
        //    STATFLAG_DEFAULT = 0,
        //    STATFLAG_NONAME = 1,
        //    STATFLAG_NOOPEN = 2
        //}

        //internal enum STGTY : int
        //{
        //    STGTY_STORAGE = 1,
        //    STGTY_STREAM = 2,
        //    STGTY_LOCKBYTES = 3,
        //    STGTY_PROPERTY = 4
        //}
        //[DllImport("ole32.dll")]
        //private static extern int StgIsStorageFile(
        //    [MarshalAs(UnmanagedType.LPWStr)] string pwcsName);
        //[DllImport("ole32.dll")]
        //private static extern int StgIsStorageILockBytes(
        //    ILockBytes plkbyt);


        //[DllImport("ole32.dll")]
        //static extern int StgOpenStorage(
        //    [MarshalAs(UnmanagedType.LPWStr)] string pwcsName,
        //    IStorage pstgPriority,
        //    STGM grfMode,
        //    IntPtr snbExclude,
        //    uint reserved,
        //    out IStorage ppstgOpen);

        //[DllImport("ole32.dll")]
        //static extern int StgOpenStorageOnILockBytes(
        //    ILockBytes plkbyt,
        //    IStorage pStgPriority,
        //    STGM grfMode,
        //    IntPtr snbEnclude,
        //    uint reserved,
        //    out IStorage ppstgOpen);
        //[DllImport("ole32.dll")]
        //static extern int CreateILockBytesOnHGlobal(
        //    IntPtr hGlobal,
        //    bool fDeleteOnRelease,
        //    out ILockBytes ppLkbyt);

        //[DllImport("ole32.dll")]
        //static extern int StgCreateDocfileOnILockBytes(ILockBytes plkbyt, STGM grfMode, int reserved, out IStorage ppstgOpen);


    }
}
