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
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Encryption
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
            if (minorVersion == 3 && majorVersion<=4)
            {
                ret = new EncryptionInfoBinary();
            }
            else if (majorVersion == 4 && minorVersion==4)
            {
                ret = new EncryptionInfoAgile();
            }
            else
            {
                throw (new NotSupportedException("Unsupported encryption format"));
            }
            ret.MajorVersion = majorVersion;
            ret.MinorVersion = minorVersion;
            ret.Read(data);
            return ret;
        }
    }
    internal enum eCipherAlgorithm
    {
        /// <summary>
        /// AES. MUST conform to the AES algorithm.
        /// </summary>
        AES,
        /// <summary>
        /// RC2. MUST conform to [RFC2268].
        /// </summary>
        RC2,
        /// <summary>
        /// RC4. 
        /// </summary>
        RC4,
        /// <summary>
        /// MUST conform to the DES algorithm.
        /// </summary>
        DES,
        /// <summary>
        /// MUST conform to the [DRAFT-DESX] algorithm.
        /// </summary>
        DESX,
        /// <summary>
        /// 3DES. MUST conform to the [RFC1851] algorithm. 
        /// </summary>
        TRIPLE_DES,
        /// 3DES_112 MUST conform to the [RFC1851] algorithm. 
        TRIPLE_DES_112        
    }
    internal enum eChainingMode
    {
        /// <summary>
        /// Cipher block chaining (CBC).
        /// </summary>
        ChainingModeCBC,
        /// <summary>
        /// Cipher feedback chaining (CFB), with 8-bit window.
        /// </summary>
        ChainingModeCFB
    }
    /// <summary>
    /// Hashalgorithm
    /// </summary>
    internal enum eHashAlogorithm
    {
        /// <summary>
        /// Sha 1-MUST conform to [RFC4634]
        /// </summary>
        SHA1,
        /// <summary>
        /// Sha 256-MUST conform to [RFC4634]
        /// </summary>
        SHA256,
        /// <summary>
        /// Sha 384-MUST conform to [RFC4634]
        /// </summary>
        SHA384,
        /// <summary>
        /// Sha 512-MUST conform to [RFC4634]
        /// </summary>
        SHA512,
        /// <summary>
        /// MD5
        /// </summary>
        MD5,
        /// <summary>
        /// MD4
        /// </summary>
        MD4,
        /// <summary>
        /// MD2
        /// </summary>
        MD2,
        /// <summary>
        /// RIPEMD-128 MUST conform to [ISO/IEC 10118]
        /// </summary>
        RIPEMD128,
        /// <summary>
        /// RIPEMD-160 MUST conform to [ISO/IEC 10118]
        /// </summary>
        RIPEMD160,
        /// <summary>
        /// WHIRLPOOL MUST conform to [ISO/IEC 10118]
        /// </summary>
        WHIRLPOOL
    }
    /// <summary>
    /// Handels the agile encryption
    /// </summary>
    internal class EncryptionInfoAgile : EncryptionInfo
    {
        XmlNamespaceManager _nsm;
        public EncryptionInfoAgile()
        {
            var nt = new NameTable();
            _nsm = new XmlNamespaceManager(nt);
            _nsm.AddNamespace("d", "http://schemas.microsoft.com/office/2006/encryption");
            _nsm.AddNamespace("c", "http://schemas.microsoft.com/office/2006/keyEncryptor/certificate");
            _nsm.AddNamespace("p", "http://schemas.microsoft.com/office/2006/keyEncryptor/password");
        }
        internal class EncryptionKeyData : XmlHelper
        {
            public EncryptionKeyData(XmlNamespaceManager nsm, XmlNode topNode) :
                base(nsm, topNode)
            {

            }
            internal byte[] SaltValue
            {
                get
                {
                    var s = GetXmlNodeString("@saltValue");
                    if (!string.IsNullOrEmpty(s))
                    {
                        return Convert.FromBase64String(s);
                    }
                    return null;
                }
                set
                {
                    SetXmlNodeString("@saltValue", Convert.ToBase64String(value));
                }
            }
            internal eHashAlogorithm HashAlgorithm
            {
                get
                {
                    return GetHashAlgorithm(GetXmlNodeString("@hashAlgorithm"));
                }
                set
                {
                    SetXmlNodeString("@hashAlgorithm", GetHashAlgorithmString(value));
                }
            }

            private eHashAlogorithm GetHashAlgorithm(string v)
            {
                switch (v)
                {
                    case "RIPEMD-128":
                        return eHashAlogorithm.RIPEMD128;
                    case "RIPEMD-160":
                        return eHashAlogorithm.RIPEMD160;
                    case "SHA-1":
                        return eHashAlogorithm.SHA1;
                    default:
                        try
                        {
                            return (eHashAlogorithm)Enum.Parse(typeof(eHashAlogorithm),v);
                        }
                        catch
                        {
                            throw (new InvalidDataException("Invalid Hash algorithm"));
                        }
                }
            }

            private string GetHashAlgorithmString(eHashAlogorithm value)
            {
                switch (value)
                {
                    case eHashAlogorithm.RIPEMD128:
                        return "RIPEMD-128";
                    case eHashAlogorithm.RIPEMD160:
                        return "RIPEMD-160";
                    case eHashAlogorithm.SHA1:
                        return "SHA-1";
                    default: 
                        return value.ToString();
                }
            }
            internal eChainingMode ChiptherChaining
            {
                get
                {
                    var v=GetXmlNodeString("@cipherChaining");
                    try
                    {
                        return (eChainingMode)Enum.Parse(typeof(eChainingMode), v);
                    }
                    catch
                    {
                        throw (new InvalidDataException("Invalid chaining mode"));
                    }
                }
                set
                {
                    SetXmlNodeString("@cipherChaining", value.ToString());
                }
            }
            internal eCipherAlgorithm CipherAlgorithm
            {
                get
                {
                    return GetCipherAlgorithm(GetXmlNodeString("@cipherAlgorithm"));
                }
                set
                {
                    SetXmlNodeString("@cipherAlgorithm", GetCipherAlgorithmString(value));
                }
            }

            private eCipherAlgorithm GetCipherAlgorithm(string v)
            {
                switch (v)
                {
                    case "3DES":
                        return eCipherAlgorithm.TRIPLE_DES;
                    case "3DES_112":
                        return eCipherAlgorithm.TRIPLE_DES_112;
                    default:
                        try
                        {
                            return (eCipherAlgorithm)Enum.Parse(typeof(eCipherAlgorithm), v);
                        }
                        catch
                        {
                            throw (new InvalidDataException("Invalid Hash algorithm"));
                        }
                }
            }

            private string GetCipherAlgorithmString(eCipherAlgorithm alg)
            {
                switch (alg)
                {
                    case eCipherAlgorithm.TRIPLE_DES:
                        return "3DES";
                    case eCipherAlgorithm.TRIPLE_DES_112:
                        return "3DES_112";                    
                    default:
                        return alg.ToString();
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
                    SetXmlNodeString("@hashSize", value.ToString());
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
                    SetXmlNodeString("@saltSize", value.ToString());
                }
            }
        }
        internal class EncryptionDataIntegrity : XmlHelper
        {
            public EncryptionDataIntegrity(XmlNamespaceManager nsm, XmlNode topNode) :
                base(nsm, topNode)
            {

            }
            internal byte[] EncryptedHmacValue
            {
                get
                {
                    var s = GetXmlNodeString("@encryptedHmacValue");
                    if (!string.IsNullOrEmpty(s))
                    {
                        return Convert.FromBase64String(s);
                    }
                    return null;
                }
                set
                {
                    SetXmlNodeString("@encryptedHmacValue", Convert.ToBase64String(value));
                }
            }
            internal byte[] EncryptedHmacKey
            {
                get
                {
                    var s = GetXmlNodeString("@encryptedHmacKey");
                    if (!string.IsNullOrEmpty(s))
                    {
                        return Convert.FromBase64String(s);
                    }
                    return null;
                }
                set
                {
                    SetXmlNodeString("@encryptedHmacKey", Convert.ToBase64String(value));
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
                    var s = GetXmlNodeString("@encryptedKeyValue");
                    if (!string.IsNullOrEmpty(s))
                    {
                        return Convert.FromBase64String(s);
                    }
                    return null;
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
                    var s = GetXmlNodeString("@encryptedVerifierHashValue");
                    if (!string.IsNullOrEmpty(s))
                    {
                        return Convert.FromBase64String(s);
                    }
                    return null;

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
                    var s = GetXmlNodeString("@encryptedVerifierHashInput");
                    if (!string.IsNullOrEmpty(s))
                    {
                        return Convert.FromBase64String(s);
                    }
                    return null;
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
                    SetXmlNodeString("@spinCount", value.ToString());
                }
            }
        }
        /*
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
           <encryption xmlns="http://schemas.microsoft.com/office/2006/encryption" xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password" xmlns:c="http://schemas.microsoft.com/office/2006/keyEncryptor/certificate">
               <keyData saltSize="16" blockSize="16" keyBits="256" hashSize="64" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA512" saltValue="pa+hrJ3s1zrY6hmVuSa5JQ==" />
               <dataIntegrity encryptedHmacKey="nd8i4sEKjsMjVN2gLo91oFN2e7bhMpWKDCAUBEpz4GW6NcE3hBXDobLksZvQGwLrPj0SUVzQA8VuDMyjMAfVCA==" encryptedHmacValue="O6oegHpQVz2uO7Om4oZijSi4kzLiiMZGIjfZlq/EFFO6PZbKitenBqe2or1REaxaI7gO/JmtJzZ1ViucqTaw4g==" />
               <keyEncryptors>
                   <keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
                      <p:encryptedKey spinCount="100000" saltSize="16" blockSize="16" keyBits="256" hashSize="64" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA512" saltValue="u2BNFAuHYn3M/WRja3/uPg==" encryptedVerifierHashInput="M0V+fRolJMRgFyI9w+AVxQ==" encryptedVerifierHashValue="V/6l9pFH7AaXFqEbsnFBfHe7gMOqFeRwaNMjc7D3LNdw6KgZzOOQlt5sE8/oG7GPVBDGfoQMTxjQydVPVy4qng==" encryptedKeyValue="B0/rbSQRiIKG5CQDH6AKYSybdXzxgKAfX1f+S5k7mNE=" />
                   </keyEncryptor></keyEncryptors></encryption>
        */
        
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

        internal XmlDocument Xml {get;set;}
        internal override void Read(byte[] data)
        {
            var byXml = new byte[data.Length - 8];
            Array.Copy(data, 8, byXml, 0, data.Length - 8);
            var xml = Encoding.UTF8.GetString(byXml);
            ReadFromXml(xml);
        }
        internal void ReadFromXml(string xml)
        {
            Xml = new XmlDocument();
            XmlHelper.LoadXmlSafe(Xml, xml, Encoding.UTF8);
            var node = Xml.SelectSingleNode("/d:encryption/d:keyData", _nsm);
            KeyData = new EncryptionKeyData(_nsm, node);
            node = Xml.SelectSingleNode("/d:encryption/d:dataIntegrity", _nsm);
            DataIntegrity = new EncryptionDataIntegrity(_nsm, node);
            KeyEncryptors = new List<EncryptionKeyEncryptor>();

            var list = Xml.SelectNodes("/d:encryption/d:keyEncryptors/d:keyEncryptor/p:encryptedKey", _nsm);
            if (list != null)
            {
                foreach (XmlNode n in list)
                {
                    KeyEncryptors.Add(new EncryptionKeyEncryptor(_nsm, n));
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
            MemoryStream ms = new MemoryStream();
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
}
