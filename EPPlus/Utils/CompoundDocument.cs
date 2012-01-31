using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using comTypes = System.Runtime.InteropServices.ComTypes;
using System.IO;

namespace OfficeOpenXml.Utils
{
    internal class CompoundDocument
    {
        internal class StoragePart
        {
            internal Dictionary<string, StoragePart> SubStorage = new Dictionary<string, StoragePart>();
            internal Dictionary<string, byte[]> DataStreams = new Dictionary<string, byte[]>();
        }
        internal StoragePart Storage = null;
        internal CompoundDocument(byte[] doc)
        {
            ILockBytes lb;
            var iret = CreateILockBytesOnHGlobal(IntPtr.Zero, true, out lb);

            IntPtr buffer = Marshal.AllocHGlobal(doc.Length);
            Marshal.Copy(doc, 0, buffer, doc.Length);
            UIntPtr readSize;
            lb.WriteAt(0, buffer, doc.Length, out readSize);
            Marshal.FreeHGlobal(buffer);

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
                    Storage = new StoragePart();
                    ReadParts(storage, Storage);
                    //foreach (var key in Storage.SubStorage["VBA"].DataStreams.Keys)
                    //{
                    //    File.WriteAllBytes(string.Format(@"c:\temp\vba\{0}.bin",key), Storage.SubStorage["VBA"].DataStreams[key]);
                    //}
                    Marshal.ReleaseComObject(storage);
                }
            }
            else
            {
                throw (new InvalidDataException(string.Format("Part is not a compound document")));
            }
        }
        #region  Compession
        internal static byte[] DecompressPart(byte[] part)
        {
            return DecompressPart(part, 0);
        }
        /// <summary>
        /// Decompression using a run length encoding algorithm.
        /// See MS-OVBA Section 2.4
        /// </summary>
        /// <param name="part"></param>
        /// <param name="startPos"></param>
        /// <returns></returns>
        internal static byte[] DecompressPart(byte[] part, int startPos)
        {

            if (part[startPos] != 1)
            {
                return null;
            }
            MemoryStream ms = new MemoryStream(4096);
            BinaryWriter br=new BinaryWriter(ms);
            int compressPos = startPos+1;
            while(compressPos<part.Length-1)
            {
                byte[] chunk = GetChunk(part, ref compressPos);
                if (chunk != null)
                {
                    br.Write(chunk);
                }
            }
            br.Flush();
            return ms.ToArray();
        }
        private static byte[] GetChunk(byte[] compBuffer, ref int pos)
        {
            ushort header = BitConverter.ToUInt16(compBuffer, pos);
            int  decomprPos=0;
            byte[] buffer = new byte[4098];
            int size = (int)(header & 0xFFF)+3;
            int endPos = pos+size;
            int a = (int)(header & 0x7000) >> 12;
            int b = (int)(header & 0x8000) >> 15;
            pos += 2;
            if (b == 1) //Compressed chunk
            {
                while (pos < compBuffer.Length && pos < endPos)
                {
                    //Decompress token
                    byte token = compBuffer[pos++];
                    if (pos >= endPos)
                        break;
                    for (int i = 0; i < 8; i++)
                    {
                        //Literal token
                        if ((token & (1 << i)) == 0)
                        {
                            buffer[decomprPos++] = compBuffer[pos++];
                        }
                        else //copy token
                        {
                            var t = BitConverter.ToUInt16(compBuffer, pos);
                            int bitCount = GetLengthBits(decomprPos);
                            int bits = (16 - bitCount);
                            ushort lengthMask = (ushort)((0xFFFF) >> bits);
                            UInt16 offsetMask = (ushort)~lengthMask;
                            var length = (lengthMask & t) + 3;
                            var offset = (offsetMask & t) >> (bitCount);
                            int source = decomprPos - offset - 1;
                            for (int c = 0; c < length; c++)
                            {
                                buffer[decomprPos++] = buffer[source++];
                            }

                            pos += 2;

                        }
                        if (pos >= endPos)
                            break;
                    }
                }
                if (decomprPos > 0)
                {
                    byte[] ret = new byte[decomprPos];
                    Array.Copy(buffer, ret, decomprPos);
                    return ret;
                }
                else
                {
                    return null;
                }
            }
            else //Raw chunk
            {
                byte[] ret = new byte[size];
                Array.Copy(compBuffer, pos, ret,0, size);
                pos += size;
                return ret;
            }
        }
        private static int GetLengthBits(int decompPos)
        {
            if (decompPos <= 16)
            {
                return 12;
            }
            else if (decompPos <= 32)
            {
                return 11;
            }
            else if (decompPos <= 64)
            {
                return 10;
            }
            else if (decompPos <= 128)
            {
                return 9;
            }
            else if (decompPos <= 256)
            {
                return 8;
            }
            else if (decompPos <= 512)
            {
                return 7;
            }
            else if (decompPos <= 1024)
            {
                return 6;
            }
            else if (decompPos <= 2048)
            {
                return 5;
            }
            else if (decompPos <= 4096)
            {
                return 4;
            }
            else
            {
                //We should never end up here, but if so this is the formula to calculate the bits...
                return 12 - (int)Math.Truncate(Math.Log(decompPos - 1 >> 4, 2) + 1);
            }
        }
        #endregion
        #region "API declare"
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
                /* [in] */ STGM grfMode,
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
        #endregion
        internal static int IsStorageFile(string Name)
        {
            return StgIsStorageFile(Name);
        }
        internal static int IsStorageILockBytes(ILockBytes lb)
        {
            return StgIsStorageILockBytes(lb);
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
        private MemoryStream ReadParts(IStorage storage, StoragePart storagePart)
        {
            MemoryStream ret = null;
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
            while (res != 1)
            {
                foreach (var item in regelt)
                {
                    if (item.type == 1)
                    {
                        IStorage subStorage;
                        storage.OpenStorage(item.pwcsName, null, STGM.DIRECT | STGM.READ | STGM.SHARE_EXCLUSIVE, IntPtr.Zero, 0, out subStorage);
                        StoragePart subStoragePart=new StoragePart();
                        storagePart.SubStorage.Add(item.pwcsName, subStoragePart);
                        ReadParts(subStorage, subStoragePart);
                    }
                    else
                    {
                        storagePart.DataStreams.Add(item.pwcsName, GetOleStream(storage, item));                    
                    }
                }
                res = pIEnumStatStg.Next(1, regelt, out fetched);
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
        //}    }
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
    }
}