/*******************************************************************************
 * You may amend and distribute as you like, but don't remove this header!
 * 
 * All rights reserved.
 * 
 * EPPlus is an Open Source project provided under the 
 * GNU General Public License (GPL) as published by the 
 * Free Software Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA 02111-1307 USA
 * 
 * See http://epplus.codeplex.com/ for details
 * 
 * The GNU General Public License can be viewed at http://www.opensource.org/licenses/gpl-license.php
 * If you unfamiliar with this license or have questions about it, here is an http://www.gnu.org/licenses/gpl-faq.html
 * 
 * The code for this project may be used and redistributed by any means PROVIDING it is 
 * not sold for profit without the author's written consent, and providing that this notice 
 * and the author's name and all copyright notices remain intact.
 * 
 * All code and executables are provided "as is" with no warranty either express or implied. 
 * The author accepts no liability for any damage or loss of business that this product may cause.
 *
 ***************************************************************************************
 * This class is created with the help of the MS-OFFCRYPTO PDF documentation... http://msdn.microsoft.com/en-us/library/cc313071(office.12).aspx
 * Decrypytion library for Office Open XML files(Lyquidity) and Sminks very nice example 
 * on "Reading compound documents in c#" on Stackoverflow. Many thanks!
 ***************************************************************************************
 *  
 * Code change notes:
 * 
 * Author							Change						Date
 *******************************************************************************
 * Jan Källman		Added		10-AUG-2010
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Text;
using System; 
using System.Runtime.InteropServices; 
using System.Runtime.InteropServices.ComTypes;
using System.IO;
using System.Security.Cryptography;
using System.IO.Packaging; 

namespace OfficeOpenXml
{
    internal class EncryptionInfo
    {
        internal short MajorVerion;
        internal short MinorVersion;
        internal Flags Flags;
        internal uint HeaderSize;
        internal EncryptionHeader Header;
        internal EncryptionVerifier Verifier;
    }
    internal class EncryptionHeader
    {
        internal Flags Flags;
        internal int SizeExtra; //MUST be 0x00000000.
        internal AlgorithmID AlgID;      //MUST be 0x0000660E (AES-128), 0x0000660F (AES-192), or 0x00006610 (AES-256).
        internal AlgorithmHashID AlgIDHash;  //MUST be 0x00008004 (SHA-1).
        internal int KeySize;    //MUST be 0x00000080 (AES-128), 0x000000C0 (AES-192), or 0x00000100 (AES-256).
        internal ProviderType ProviderType;    //SHOULD<10> be 0x00000018 (AES).
        internal int Reserved1;      //Undefined and MUST be ignored.
        internal int Reserved2;      //MUST be 0x00000000 and MUST be ignored.
        internal string CSPName;     //SHOULD<11> be set to either "Microsoft Enhanced RSA and AES Cryptographic Provider" or "Microsoft Enhanced RSA and AES Cryptographic Provider (Prototype)" as a null-terminated Unicode string.
    }
    internal class EncryptionVerifier
    {
        internal uint SaltSize;              // An unsigned integer that specifies the size of the Salt field. It MUST be 0x00000010.
        internal byte[] Salt;                //(16 bytes): An array of bytes that specifies the salt value used during password hash generation. It MUST NOT be the same data used for the verifier stored encrypted in the EncryptedVerifier field.
        internal byte[] EncryptedVerifier;   //(16 bytes): MUST be the randomly generated Verifier value encrypted using the algorithm chosen by the implementation.
        internal uint VerifierHashSize;      //(4 bytes): An unsigned integer that specifies the number of bytes needed to contain the hash of the data used to generate the EncryptedVerifier field.
        internal byte[] EncryptedVerifierHash; //(variable): An array of bytes that contains the encrypted form of the hash of the randomly generated Verifier value. The length of the array MUST be the size of the encryption block size multiplied by the number of blocks needed to encrypt the hash of the Verifier. If the encryption algorithm is RC4, the length MUST be 20 bytes. If the encryption algorithm is AES, the length MUST be 32 bytes.
    }
    [Flags]
    internal enum Flags
    {
        Reserved1 = 1,   // (1 bit): MUST be set to zero, and MUST be ignored.
        Reserved2 = 2,   // (1 bit): MUST be set to zero, and MUST be ignored.
        fCryptoAPI= 4,  // (1 bit): A flag that specifies whether CryptoAPI RC4 or [ECMA-376] encryption is used. It MUST be set to 1 unless fExternal is 1. If fExternal is set to 1, it MUST be set to zero.        
        fDocProps = 8,   // (1 bit): MUST be set to zero if document properties are encrypted. Otherwise, it MUST be set to 1. Encryption of document properties is specified in section 2.3.5.4.
        fExternal = 16,   // (1 bit): If extensible encryption is used, it MUST be set to 1. Otherwise, it MUST be set to zero. If this field is set to 1, all other fields in this structure MUST be set to zero.
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
            uint grfMode,
            IntPtr snbEnclude,
            uint reserved,
            out IStorage ppstgOpen);
        [DllImport("ole32.dll")]
        static extern int CreateILockBytesOnHGlobal(
            IntPtr hGlobal,
            bool fDeleteOnRelease,
            out ILockBytes ppLkbyt);

        /// <summary>
        /// Read the package from the OLE document and decrypt it using the supplied password
        /// </summary>
        /// <param name="fi">The file</param>
        /// <param name="password"></param>
        /// <returns></returns>
        public MemoryStream GetStream(FileInfo fi, string password)
        {
            MemoryStream ret = null;
            if (StgIsStorageFile(fi.FullName) == 0)
            {
                IStorage storage = null;
                //IStorage pIChildStorage;
                if (StgOpenStorage(
                    fi.FullName,
                    null,
                    STGM.DIRECT | STGM.READ | STGM.SHARE_EXCLUSIVE,
                    IntPtr.Zero,
                    0,
                    out storage) == 0)
                {
                    ret = GetStreamFromPackage(password, storage);
                }
            }
            return ret;
        }
        private MemoryStream GetStreamFromPackage(string password, IStorage storage)
        {
            MemoryStream ret=null;        
            System.Runtime.InteropServices.ComTypes.STATSTG statstg;

            storage.Stat(out statstg, (uint)STATFLAG.STATFLAG_DEFAULT);

            IEnumSTATSTG pIEnumStatStg = null;
            storage.EnumElements(0, IntPtr.Zero, 0, out pIEnumStatStg);

            System.Runtime.InteropServices.ComTypes.STATSTG[] regelt = { statstg };
            uint fetched = 0;
            uint res = pIEnumStatStg.Next(1, regelt, out fetched);

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
                            encryptionInfo = GetEncryptionInfo(data);
                            break;
                        case "EncryptedPackage":
                            data = GetOleStream(storage, statstg);
                            ret = DecryptDocument(data, encryptionInfo, password);
                            break;
                    }

                    if ((res = pIEnumStatStg.Next(1, regelt, out fetched)) != 1)
                    {
                        statstg = regelt[0];
                    }
                }
            }
            return ret;
        }
        private MemoryStream DecryptDocument(byte[] data, EncryptionInfo encryptionInfo, string password)
        {
            if (encryptionInfo == null)
            {
                throw(new Exception("Invalid document. EncryptionInfo is missing"));
            }
            long size = BitConverter.ToInt64(data,0);

            var encryptedData = new byte[data.Length - 8];
            Array.Copy(data, 8, encryptedData, 0, encryptedData.Length);
            
            MemoryStream doc = new MemoryStream();

            if (encryptionInfo.Header.AlgID == AlgorithmID.AES128 || (encryptionInfo.Header.AlgID == AlgorithmID.Flags  && ((encryptionInfo.Flags & (Flags.fAES | Flags.fExternal | Flags.fCryptoAPI)) == (Flags.fAES | Flags.fCryptoAPI)))
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

                var key=GetPasswordHash(password, encryptionInfo);

                ICryptoTransform decryptor = decryptKey.CreateDecryptor(
                                                         key,
                                                         null);


                MemoryStream dataStream = new MemoryStream(encryptedData);

                CryptoStream cryptoStream = new CryptoStream(dataStream,
                                                              decryptor,
                                                              CryptoStreamMode.Read);

                var decryptedData = new byte[size];

                cryptoStream.Read(decryptedData,0,(int)size);
                doc.Write(decryptedData, 0, (int)size);
            }
            return doc;
        }
        private byte[] GetOleStream(IStorage storage, System.Runtime.InteropServices.ComTypes.STATSTG statstg)
        {
            IStream pIStream;
            storage.OpenStream(statstg.pwcsName,
               IntPtr.Zero,
               (uint)(STGM.READ | STGM.SHARE_EXCLUSIVE),
               0,
               out pIStream);

            byte[] data = new byte[statstg.cbSize];
            pIStream.Read(data, (int)statstg.cbSize, IntPtr.Zero);
            return data;
        } 
        private EncryptionInfo GetEncryptionInfo(byte[] data)
        {
            EncryptionInfo info =new EncryptionInfo();

            info.MajorVerion = BitConverter.ToInt16(data,0);
            info.MinorVersion = BitConverter.ToInt16(data,2);

            info.Flags = (Flags)BitConverter.ToInt32(data, 4);
            info.HeaderSize = (uint)BitConverter.ToInt32(data, 8);
        
            /**** EncryptionHeader ****/
            info.Header = new EncryptionHeader();
            info.Header.Flags = (Flags)BitConverter.ToInt32(data, 12);
            info.Header.SizeExtra = BitConverter.ToInt32(data, 16);
            info.Header.AlgID =(AlgorithmID) BitConverter.ToInt32(data, 20);
            info.Header.AlgIDHash = (AlgorithmHashID)BitConverter.ToInt32(data, 24);
            info.Header.KeySize = BitConverter.ToInt32(data, 28);
            info.Header.ProviderType = (ProviderType)BitConverter.ToInt32(data, 32);
            info.Header.Reserved1 = BitConverter.ToInt32(data, 36);
            info.Header.Reserved2 = BitConverter.ToInt32(data, 40);

            byte[] text = new byte[(int)info.HeaderSize - 34];
            Array.Copy(data, 44, text, 0, (int)info.HeaderSize - 34);
            info.Header.CSPName = UTF8Encoding.Unicode.GetString(text);

            int pos = (int)info.HeaderSize + 12;

            /**** EncryptionVerifier ****/
            info.Verifier = new EncryptionVerifier();
            info.Verifier.SaltSize = (uint)BitConverter.ToInt32(data, pos);
            info.Verifier.Salt = new byte[info.Verifier.SaltSize];
            
            Array.Copy(data, pos + 4, info.Verifier.Salt, 0, info.Verifier.SaltSize);
            
            info.Verifier.EncryptedVerifier = new byte[16];
            Array.Copy(data, pos + 20, info.Verifier.EncryptedVerifier, 0, 16);

            info.Verifier.VerifierHashSize = (uint)BitConverter.ToInt32(data, pos+36);
            info.Verifier.EncryptedVerifierHash = new byte[info.Verifier.VerifierHashSize];
            Array.Copy(data, pos + 40, info.Verifier.EncryptedVerifierHash, 0, info.Verifier.VerifierHashSize);

            return info;
        }
        /// <summary>
        /// Create the hash.
        /// This method is written with the help of Lyquidity library, many thanks for this nice sample
        /// </summary>
        /// <param name="password">The password</param>
        /// <param name="encryptionInfo">The encryption info extracted from the ENCRYPTIOINFO stream inside the OLE document</param>
        /// <returns>The hash to encrypt the document</returns>
        private byte[] GetPasswordHash(string password, EncryptionInfo encryptionInfo)
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
                    throw new Exception("RC4 Hash provider is not supported. Must be SHA1(AlgIDHash == 0x8004)");
                }
                else
                {
                    throw new Exception("Hash provider is invalid. Must be SHA1(AlgIDHash == 0x8004)");
                }

                hash = hashProvider.ComputeHash(CombinePassword(encryptionInfo.Verifier.Salt, password));

                //Iterate 50 000 times, inserting i in first 4 bytes and then the prev. hash in byte 5-24
                for (int i = 0; i < 50000; i++)
                {
                    Array.Copy(BitConverter.GetBytes(i), tempHash, 4);
                    Array.Copy(hash, 0, tempHash, 4, hash.Length);     
               
                    hash = hashProvider.ComputeHash(tempHash);
                }

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
                if (encryptionInfo.Verifier.VerifierHashSize > keySizeBytes)
                    return FixHashSize(X1,keySizeBytes);

                //Else XOR hash bytes with 0x5C and fill the rest with 0x5C
                for (int i = 0; i < derivedKey.Length; i++)
                    derivedKey[i] = (byte)(i < hash.Length ? 0x5C ^ hash[i] : 0x5C);

                byte[] X2 = hashProvider.ComputeHash(derivedKey);

                //Join the two and return 
                byte[] join = new byte[X1.Length + X2.Length];

                Array.Copy(X1, 0, join, 0, X1.Length);
                Array.Copy(X2, 0, join, X1.Length, X2.Length);

                return FixHashSize(join,keySizeBytes);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            return null;
        }

        private byte[] FixHashSize(byte[] hash, int size)
        {
            byte[] buff = new byte[size];
            Array.Copy(hash, buff, size);
            return buff;
        }
        private static byte[] CombinePassword(byte[] salt, string password)
        {
            if (password == "")
            {
                password = "VelvetSweatshop";   //Used if Password if blank
            }
            // Convert password to unicode...
            byte[] passwordBuf = UnicodeEncoding.Unicode.GetBytes(password);
            
            byte[] inputBuf = new byte[salt.Length + passwordBuf.Length];
            Array.Copy(salt, inputBuf, salt.Length);
            Array.Copy(passwordBuf, 0, inputBuf, salt.Length, passwordBuf.Length);            
            return inputBuf;
        }
    }
    [ComImport] 
    [Guid("0000000d-0000-0000-C000-000000000046")] 
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)] 
    public interface IEnumSTATSTG 
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
            /* [out] */ out IStream ppstm); 
        
        void OpenStream( 
            /* [string][in] */ string pwcsName, 
            /* [unique][in] */ IntPtr reserved1, 
            /* [in] */ uint grfMode, 
            /* [in] */ uint reserved2, 
            /* [out] */ out IStream ppstm); 
 
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
            /* [in] */ uint ciidExclude, 
            /* [size_is][unique][in] */ Guid rgiidExclude, // should this be an array? 
            /* [unique][in] */ IntPtr snbExclude, 
            /* [unique][in] */ IStorage pstgDest); 
 
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
    public interface ILockBytes
    {
        void ReadAt(long ulOffset, System.IntPtr pv, int cb, out UIntPtr pcbRead);
        void WriteAt(long ulOffset, System.IntPtr pv, int cb, out UIntPtr pcbWritten);
        void Flush();
        void SetSize(long cb);
        void LockRegion(long libOffset, long cb, int dwLockType);
        void UnlockRegion(long libOffset, long cb, int dwLockType);
        void Stat(out System.Runtime.InteropServices.STATSTG pstatstg, int grfStatFlag);
    }
    [Flags] 
    public enum STGM : int 
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
 
    public enum STATFLAG : uint 
    { 
        STATFLAG_DEFAULT = 0, 
        STATFLAG_NONAME = 1, 
        STATFLAG_NOOPEN = 2 
    } 
 
    public enum STGTY : int 
    { 
        STGTY_STORAGE = 1, 
        STGTY_STREAM = 2, 
        STGTY_LOCKBYTES = 3, 
        STGTY_PROPERTY = 4 
    }
}
