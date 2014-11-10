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
 *******************************************************************************
 * Jan Källman		Added		01-01-2012
 * Jan Källman      Added compression support 27-03-2012
 *******************************************************************************/
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
            public StoragePart()
            {

            }
            internal Dictionary<string, StoragePart> SubStorage = new Dictionary<string, StoragePart>();
            internal Dictionary<string, byte[]> DataStreams = new Dictionary<string, byte[]>();
        }
        internal StoragePart Storage = null;
        internal CompoundDocument()
        {
            Storage = new CompoundDocument.StoragePart();
        }
        internal CompoundDocument(FileInfo fi)
        {
            Read(fi);
        }
        internal CompoundDocument(ILockBytes lb)
        {
            Read(lb);
        }
        internal CompoundDocument(byte[] doc)
        {
            Read(doc);
        }
        internal void Read(FileInfo fi)
        {
            var b = File.ReadAllBytes(fi.FullName);
            Read(b);
        }
        internal void Read(byte[] doc)
        {
            ILockBytes lb;
            var iret = CreateILockBytesOnHGlobal(IntPtr.Zero, true, out lb);

            IntPtr buffer = Marshal.AllocHGlobal(doc.Length);
            Marshal.Copy(doc, 0, buffer, doc.Length);
            UIntPtr readSize;
            lb.WriteAt(0, buffer, doc.Length, out readSize);
            Marshal.FreeHGlobal(buffer);

            Read(lb);
        }

        internal void Read(ILockBytes lb)
        {
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
                    Marshal.ReleaseComObject(storage);
                }
            }
            else
            {
                throw (new InvalidDataException(string.Format("Part is not a compound document")));
            }
        }
        #region  Compression
        /// <summary>
        /// Compression using a run length encoding algorithm.
        /// See MS-OVBA Section 2.4
        /// </summary>
        /// <param name="part">Byte array to decompress</param>
        /// <returns></returns>
        internal static byte[] CompressPart(byte[] part)
        {
            MemoryStream ms = new MemoryStream(4096);
            BinaryWriter br = new BinaryWriter(ms);
            br.Write((byte)1);

            int compStart = 1;
            int compEnd = 4098;
            int decompStart = 0;
            int decompEnd = part.Length < 4096 ? part.Length : 4096;

            while (decompStart < decompEnd && compStart < compEnd)
            {
                byte[] chunk = CompressChunk(part, ref decompStart);
                ushort header;
                if (chunk == null || chunk.Length == 0)
                {
                    header = 4096 | 0x600;  //B=011 A=0
                }
                else
                {
                    header = (ushort)(((chunk.Length - 1) & 0xFFF));
                    header |= 0xB000;   //B=011 A=1
                    br.Write(header);
                    br.Write(chunk);                    
                }
                decompEnd = part.Length < decompStart + 4096 ? part.Length : decompStart+4096;
            }

            
            br.Flush();
            return ms.ToArray();        
        }
        private static byte[] CompressChunk(byte[] buffer, ref int startPos)
        {
            var comprBuffer=new byte[4096];
            int flagPos = 0;
            int cPos=1;
            int dPos = startPos;
            int dEnd=startPos+4096 < buffer.Length? startPos+4096 : buffer.Length;
            while(dPos<dEnd)
            {
                byte tokenFlags = 0;
                for (int i = 0; i < 8; i++)
                {
                    if (dPos - startPos > 0)
                    {
                        int bestCandidate = -1;
                        int bestLength = 0;
                        int candidate = dPos - 1;
                        int bitCount = GetLengthBits(dPos-startPos);
                        int bits = (16 - bitCount);
                        ushort lengthMask = (ushort)((0xFFFF) >> bits);

                        while (candidate >= startPos)
                        {
                            if (buffer[candidate] == buffer[dPos])
                            {
                                int length = 1;

                                while (buffer.Length > dPos + length && buffer[candidate + length] == buffer[dPos + length] && length < lengthMask && dPos+length < dEnd)
                                {
                                    length++;
                                }
                                if (length > bestLength)
                                {
                                    bestCandidate = candidate;
                                    bestLength = length;
                                    if (bestLength == lengthMask)
                                    {
                                        break;
                                    }
                                }
                            }
                            candidate--;
                        }
                        if (bestLength >= 3)    //Copy token
                        {
                            tokenFlags |= (byte)(1 << i);

                            UInt16 offsetMask = (ushort)~lengthMask;
                            ushort token = (ushort)(((ushort)(dPos - (bestCandidate+1))) << (bitCount) | (ushort)(bestLength - 3));
                            Array.Copy(BitConverter.GetBytes(token), 0, comprBuffer, cPos, 2);
                            dPos = dPos + bestLength;
                            cPos += 2;
                            //SetCopy Token                        
                        }
                        else
                        {
                            comprBuffer[cPos++] = buffer[dPos++];
                        }
                    }
                    
                    else
                    {
                        comprBuffer[cPos++] = buffer[dPos++];
                    }
                    if (dPos >= dEnd) break;
                }
                comprBuffer[flagPos] = tokenFlags;
                flagPos = cPos++;
            }
            var ret = new byte[cPos - 1];
            Array.Copy(comprBuffer, ret, ret.Length);
            startPos = dEnd;
            return ret;
        }
        internal static byte[] DecompressPart(byte[] part)
        {
            return DecompressPart(part, 0);
        }
        /// <summary>
        /// Decompression using a run length encoding algorithm.
        /// See MS-OVBA Section 2.4
        /// </summary>
        /// <param name="part">Byte array to decompress</param>
        /// <param name="startPos"></param>
        /// <returns></returns>
        internal static byte[] DecompressPart(byte[] part, int startPos)
        {

            if (part[startPos] != 1)
            {
                return null;
            }
            MemoryStream ms = new MemoryStream(4096);
            int compressPos = startPos + 1;
            while(compressPos < part.Length-1)
            {
                DecompressChunk(ms, part, ref compressPos);
            }
            return ms.ToArray();
        }
        private static void DecompressChunk(MemoryStream ms, byte[] compBuffer, ref int pos)
        {
            ushort header = BitConverter.ToUInt16(compBuffer, pos);
            int  decomprPos=0;
            byte[] buffer = new byte[4198]; //Add an extra 100 byte. Some workbooks have overflowing worksheets.
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
                            ms.WriteByte(compBuffer[pos]);
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
                            if (decomprPos + length >= buffer.Length)
                            {
                                // Be lenient on decompression, so extend our decompression
                                // buffer. Excel generated VBA projects do encounter this issue.
                                // One would think (not surprisingly that the VBA project spec)
                                // over emphasizes the size restrictions of a DecompressionChunk.
                                var largerBuffer = new byte[buffer.Length + 4098];
                                Array.Copy(buffer, largerBuffer, decomprPos);
                                buffer = largerBuffer;
                            }
                            ms.Write(buffer, source, length);
                            // Even though we've written to the MemoryStream,
                            // We still should decompress the token into this buffer
                            // in case a later token needs to use the bytes we're
                            // about to decompress.
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
                    ms.Write(buffer, 0, decomprPos);
                    return;
                }
                else
                {
                    return;
                }
            }
            else //Raw chunk
            {
                ms.Write(compBuffer, pos, size);
                pos += size;
                return;
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
        internal static ILockBytes GetLockbyte(MemoryStream stream)
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
        internal byte[] Save()
        {
            ILockBytes lb;
            var iret = CreateILockBytesOnHGlobal(IntPtr.Zero, true, out lb);

            IStorage storage = null;
            byte[] ret = null;

            //Create the document in-memory
            if (StgCreateDocfileOnILockBytes(lb,
                    STGM.CREATE | STGM.READWRITE | STGM.SHARE_EXCLUSIVE | STGM.TRANSACTED, 
                    0,
                    out storage)==0)
            {
                foreach(var store in this.Storage.SubStorage)
                {
                    CreateStore(store.Key, store.Value, storage);
                }
                CreateStreams(this.Storage, storage);                                
                lb.Flush();
                
                //Now copy the unmanaged stream to a byte array --> memory stream
                var statstg = new comTypes.STATSTG();
                lb.Stat(out statstg, 0);
                int size = (int)statstg.cbSize;
                IntPtr buffer = Marshal.AllocHGlobal(size);
                UIntPtr readSize;
                ret=new byte[size];
                lb.ReadAt(0, buffer, size, out readSize);
                Marshal.Copy(buffer, ret, 0, size);
                Marshal.FreeHGlobal(buffer);
            }
            Marshal.ReleaseComObject(storage);
            Marshal.ReleaseComObject(lb);

            return ret;
        }

        private void CreateStore(string name, StoragePart subStore, IStorage storage)
        {
            IStorage subStorage;
            storage.CreateStorage(name, (uint)(STGM.CREATE | STGM.WRITE | STGM.DIRECT | STGM.SHARE_EXCLUSIVE), 0, 0, out subStorage);
            storage.Commit(0);
            foreach (var store in subStore.SubStorage)
            {
                CreateStore(store.Key, store.Value, subStorage);
            }
            
            CreateStreams(subStore, subStorage);
        }

        private void CreateStreams(StoragePart subStore, IStorage subStorage)
        {
            foreach (var ds in subStore.DataStreams)
            {
                comTypes.IStream stream;
                subStorage.CreateStream(ds.Key, (uint)(STGM.CREATE | STGM.WRITE | STGM.DIRECT | STGM.SHARE_EXCLUSIVE), 0, 0, out stream);
                stream.Write(ds.Value, ds.Value.Length, IntPtr.Zero);
            }
            subStorage.Commit(0);
        }
    }
}