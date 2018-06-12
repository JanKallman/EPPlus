/*******************************************************************************
 * You may amend and distribute as you like, but don't remove this header!
 *
 * EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
 * See https://github.com/JanKallman/EPPlus for details.
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
 * Jan Källman		Added		28-MAR-2017
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
namespace OfficeOpenXml.Utils.CompundDocument
{
    /// <summary>
    /// Read and write a compound document.
    /// Read spec here https://winprotocoldoc.blob.core.windows.net/productionwindowsarchives/MS-CFB/[MS-CFB].pdf
    /// </summary>
    internal class CompoundDocumentFile : IDisposable
    {
        public CompoundDocumentFile()
        {
            RootItem = new CompoundDocumentItem() { Name = "<Root>", Children=new List<CompoundDocumentItem>(), ObjectType=5 };
            minorVersion = 0x3E;
            majorVersion = 3;
            sectorShif = 9;
            minSectorShift = 6;

            _sectorSize = 1 << sectorShif;
            _miniSectorSize = 1 << minSectorShift;
            _sectorSizeInt = _sectorSize / 4;
        }
        internal CompoundDocumentFile(FileInfo fi) : this(File.ReadAllBytes(fi.FullName))
        {
            
        }
        public CompoundDocumentFile(byte[] file) : this(new MemoryStream(file))
        {
        }
        public CompoundDocumentFile(MemoryStream ms)
        {
            ms.Seek(0, SeekOrigin.Begin);   //Fixes issue #60
            Read(new BinaryReader(ms));
        }
        private struct DocWriteInfo
        {
            internal List<int> DIFAT, FAT, miniFAT;
        }
        #region Constants
        const int miniFATSectorSize = 64;
        const int FATSectorSizeV3= 512;
        const int FATSectorSizeV4 = 4096;

        const int DIFAT_SECTOR = -4; //0xFFFFFFFC;
        const int FAT_SECTOR = -3;   //0xFFFFFFFD;
        const int END_OF_CHAIN = -2; //0xFFFFFFFE;
        const int FREE_SECTOR = -1;  //0xFFFFFFFF;

        static readonly byte[] header = new byte[] { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 };
        #endregion
        #region Private Fields
        short minorVersion;
        short majorVersion;            
        int numberOfDirectorySector;
        short sectorShif, minSectorShift;       //Bits for sector size

        int _numberOfFATSectors;            // (4 bytes): This integer field contains the count of the number of FAT sectors in the compound file.
        int _firstDirectorySectorLocation;  // (4 bytes): This integer field contains the starting sector number for the directory stream.            
        int _transactionSignatureNumber;    // (4 bytes): This integer field MAY contain a sequence number that is incremented every time the compound file is saved by an implementation that supports file transactions.This is the field that MUST be set to all zeroes if file transactions are not implemented.<1> 
        int _miniStreamCutoffSize;          // (4 bytes): This integer field MUST be set to 0x00001000. This field specifies the maximum size of a user-defined data stream that is allocated from the mini FAT and mini stream, and that cutoff is 4,096 bytes.Any user-defined data stream that is larger than or equal to this cutoff size must be allocated as normal sectors from the FAT.
        int _firstMiniFATSectorLocation;    // (4 bytes): This integer field contains the starting sector number for the mini FAT. 
        int _numberofMiniFATSectors;        // (4 bytes): This integer field contains the count of the number of mini FAT sectors in the compound file. 
        int _firstDIFATSectorLocation;      // (4 bytes): This integer field contains the starting sector number for the DIFAT. 
        int _numberofDIFATSectors;          // (4 bytes): This integer field contains the count of the number of DIFAT sectors in the compound file. 

        List<byte[]> _sectors, _miniSectors;
        int _sectorSize, _miniSectorSize;
        int _sectorSizeInt;
        int _currentDIFATSectorPos, _currentFATSectorPos, _currentDirSectorPos;
        int _prevDirFATSectorPos;

        #endregion
        public CompoundDocumentItem RootItem { get; set; }
        /// <summary>
        /// Verifies that the header is correct.
        /// </summary>
        /// <param name="fi">The file</param>
        /// <returns></returns>
        public static bool IsCompoundDocument(FileInfo fi)
        {
            try
            {
                var fs = fi.OpenRead();
                var b = new byte[8];
                fs.Read(b, 0, 8);
                return IsCompoundDocument(b);
            }
            catch
            {
                return false;
            }            
        }
        public static bool IsCompoundDocument(MemoryStream ms)
        {
            var pos = ms.Position;
            ms.Position = 0;
            var b=new byte[8];
            ms.Read(b, 0, 8);
            ms.Position = pos;
            return IsCompoundDocument(b);
        }
        public static bool IsCompoundDocument(byte[] b)
        {
            if (b==null || b.Length < 8) return false;
            for (int i = 0; i < 8; i++)
            {
                if (b[i] != header[i])
                {
                    return false;
                }
            }
            return true;
        }
    #region Read
    internal void Read(BinaryReader br)
        {
            br.ReadBytes(8);    //Read header
            br.ReadBytes(16);   //Header CLSID (16 bytes): Reserved and unused class ID that MUST be set to all zeroes (CLSID_NULL). 
            minorVersion = br.ReadInt16();
            majorVersion = br.ReadInt16();
            br.ReadInt16(); //Byte order
            sectorShif = br.ReadInt16();
            minSectorShift = br.ReadInt16();

            _sectorSize = 1 << sectorShif;
            _miniSectorSize = 1 << minSectorShift;
            _sectorSizeInt = _sectorSize / 4;
            br.ReadBytes(6);    //Reserved
            numberOfDirectorySector = br.ReadInt32();
            _numberOfFATSectors = br.ReadInt32();
            _firstDirectorySectorLocation = br.ReadInt32();
            _transactionSignatureNumber = br.ReadInt32();
            _miniStreamCutoffSize = br.ReadInt32();
            _firstMiniFATSectorLocation = br.ReadInt32();
            _numberofMiniFATSectors = br.ReadInt32();
            _firstDIFATSectorLocation = br.ReadInt32();
            _numberofDIFATSectors = br.ReadInt32();
            var dwi = new DocWriteInfo() { DIFAT = new List<int>(), FAT = new List<int>(), miniFAT = new List<int>() };

            for (int i = 0; i < 109; i++)
            {
                var d = br.ReadInt32();
                if (d >= 0)
                {   
                    dwi.DIFAT.Add(d);
                }
            }

            LoadSectors(br);
            if (_firstDIFATSectorLocation > 0)
            {
                LoadDIFATSectors(dwi);
            }

            dwi.FAT = ReadFAT(_sectors, dwi);
            var dir = ReadDirectories(_sectors, dwi);

            LoadMinSectors(ref dwi, dir);
            foreach (var d in dir)
            {
                if (d.Stream == null && d.StreamSize > 0)
                {
                    if (d.StreamSize < _miniStreamCutoffSize)
                    {
                        d.Stream = GetStream(d.StartingSectorLocation, d.StreamSize, dwi.miniFAT, _miniSectors);
                    }
                    else
                    {
                        d.Stream = GetStream(d.StartingSectorLocation, d.StreamSize, dwi.FAT, _sectors);
                    }
                }
            }
            AddChildTree(dir[0], dir);
        }
        private void LoadDIFATSectors(DocWriteInfo dwi)
        {
            var nextSector = _firstDIFATSectorLocation;
            while (nextSector > 0)
            {
                var brDI = new BinaryReader(new MemoryStream(_sectors[nextSector]));
                var sect = -1;
                while (brDI.BaseStream.Position < _sectorSize)
                {
                    if (sect > 0)
                    {
                        dwi.DIFAT.Add(sect);
                    }
                    sect = brDI.ReadInt32();
                }
                nextSector = sect;
            }
        }
        private void LoadSectors(BinaryReader br)
        {
            _sectors = new List<byte[]>();
            while (br.BaseStream.Position < br.BaseStream.Length)
            {
                _sectors.Add(br.ReadBytes(_sectorSize));
            }
        }
        private void LoadMinSectors(ref DocWriteInfo dwi, List<CompoundDocumentItem> dir)
        {
            dwi.miniFAT = ReadMiniFAT(_sectors,dwi);
            dir[0].Stream = GetStream(dir[0].StartingSectorLocation, dir[0].StreamSize, dwi.FAT, _sectors);
            GetMiniSectors(dir[0].Stream);
        }
        private void GetMiniSectors(byte[] miniFATStream)
        {
            var br = new BinaryReader(new MemoryStream(miniFATStream));
            _miniSectors = new List<byte[]>();
            while (br.BaseStream.Position < br.BaseStream.Length)
            {
                _miniSectors.Add(br.ReadBytes(_miniSectorSize));
            }
        }
        private byte[] GetStream(int startingSectorLocation, long streamSize, List<int> FAT, List<byte[]> sectors)
        {
            var ms = new MemoryStream();
            var bw = new BinaryWriter(ms);

            var size = 0;
            var nextSector = startingSectorLocation;
            while(size<streamSize)
            {
                if (streamSize > size + sectors[nextSector].Length)
                {
                    bw.Write(sectors[nextSector]);
                    size += sectors[nextSector].Length;
                }
                else
                {                        
                    var part= new byte[streamSize-size];
                    Array.Copy(sectors[nextSector], part, (int)streamSize - size);
                    bw.Write(part);
                    size += part.Length;
                }
                nextSector = FAT[nextSector];
            }
            bw.Flush();
            return ms.ToArray();
        }
        private List<int> ReadMiniFAT(List<byte[]> sectors, DocWriteInfo dwi)
        {
            var l = new List<int>();
            var nextSector = _firstMiniFATSectorLocation;
            while(nextSector!=END_OF_CHAIN)
            {
                var br = new BinaryReader(new MemoryStream(sectors[nextSector]));
                while (br.BaseStream.Position < _sectorSize)
                {
                    var d = br.ReadInt32();
                    l.Add(d);
                }
                nextSector = dwi.FAT[nextSector];
            }
            return l;
        }
        private List<CompoundDocumentItem> ReadDirectories(List<byte[]> sectors, DocWriteInfo dwi)
        {
            var dir = new List<CompoundDocumentItem>();
            var nextSector = _firstDirectorySectorLocation;
            while (nextSector != END_OF_CHAIN)
            {
                ReadDirectory(sectors, nextSector, dir);
                nextSector = dwi.FAT[nextSector];
            }
            return dir;
        }
        private List<int> ReadFAT(List<byte[]> sectors, DocWriteInfo dwi)
        {
            var l = new List<int>();
            foreach (var i in dwi.DIFAT)
            {
                var br = new BinaryReader(new MemoryStream(sectors[i]));
                while (br.BaseStream.Position < _sectorSize)
                {
                    var d = br.ReadInt32();
                    l.Add(d);
                }
            }
            return l;
        }
        private void ReadDirectory(List<byte[]> sectors, int index, List<CompoundDocumentItem> l)
        {

            var br = new BinaryReader(new MemoryStream(sectors[index]));

            while (br.BaseStream.Position < br.BaseStream.Length)
            {
                var e = new CompoundDocumentItem();
                e.Read(br);
                if (e.ObjectType!=0)
                {
                    l.Add(e);
                }
            }
        }
        internal void AddChildTree(CompoundDocumentItem e, List<CompoundDocumentItem> dirs)
        {
            if (e._handled == true) return;
            e._handled = true;
            if (e.ChildID > 0)
            {
                var c = dirs[e.ChildID];
                c.Parent = e;
                e.Children.Add(c);
                AddChildTree(c, dirs);
            }
            if (e.LeftSibling > 0)
            {
                var c = dirs[e.LeftSibling];
                c.Parent = e.Parent;
                c.Parent.Children.Insert(e.Parent.Children.IndexOf(e), c);
                AddChildTree(c, dirs);
            }
            if (e.RightSibling > 0)
            {
                var c = dirs[e.RightSibling];
                c.Parent = e.Parent;
                e.Parent.Children.Insert(e.Parent.Children.IndexOf(e) + 1, c);
                AddChildTree(c, dirs);
            }
            if (e.ObjectType == 5)
            {
                RootItem = e;
            }
        }
        internal void AddLeftSiblingTree(CompoundDocumentItem e, List<CompoundDocumentItem> dirs)
        {
            if (e.LeftSibling > 0)
            {
                var c = dirs[e.LeftSibling];
                if (c.Parent != null)
                {
                    c.Parent = e.Parent;
                    c.Parent.Children.Insert(e.Parent.Children.IndexOf(e), c);
                    e._handled = true;
                    AddLeftSiblingTree(c, dirs);
                }
            }
        }
        internal void AddRightSiblingTree(CompoundDocumentItem e, List<CompoundDocumentItem> dirs)
        {
            if (e.RightSibling > 0)
            {
                var c = dirs[e.RightSibling];
                c.Parent = e.Parent;
                e.Parent.Children.Insert(e.Parent.Children.IndexOf(e) + 1, c);
                e._handled = true;
                AddRightSiblingTree(c, dirs);
            }
        }
    #endregion
    #region Write
    public void Write(MemoryStream ms)
    {
            var bw = new BinaryWriter(ms);

            //InitValues
            minorVersion = 62;
            majorVersion = 3;
            sectorShif = 9;                 //2^9=512 bytes for version 3 documents 
            minSectorShift = 6;             //2^6=64 bytes
            _miniStreamCutoffSize = 4096;
            _transactionSignatureNumber = 0;
            _firstDIFATSectorLocation = END_OF_CHAIN;
            _firstDirectorySectorLocation = 1;
            _firstMiniFATSectorLocation = 2;
            _numberOfFATSectors = 1;

            _currentDIFATSectorPos = 76;             //DIFAT Position in the header
            _currentFATSectorPos = _sectorSize;      //First FAT sector starts at Sector 0
            _currentDirSectorPos = _sectorSize * 2;  //First FAT sector starts at Sector 1
            _prevDirFATSectorPos = _sectorSize + 4;  //Dir sector starts FAT position 1 (4 for int size)

            bw.Write(new byte[512 * 4]);            //Allocate for Header and first FAT, Directory och MiniFAT sectors
            WritePosition(bw, 0, ref _currentDIFATSectorPos, false);
            WritePosition(bw, new int[] { FAT_SECTOR, END_OF_CHAIN, END_OF_CHAIN }, ref _currentFATSectorPos);  //First sector is first FAT sector, second is First Dir sector, thirs is first Mini FAT sector.

            var dirs = FlattenDirs();

            //Write directories
            WriteDirs(bw, dirs);
                
            //Fill empty DISectors up to 109
            FillDIFAT(bw);
            //Finally write the header information in the top of the file
            WriteHeader(bw);
        }

        private List<CompoundDocumentItem> FlattenDirs()
        {
            var l = new List<CompoundDocumentItem>();
            InitItem(RootItem);
            l.Add(RootItem);
            RootItem.ChildID = AddChildren(RootItem, l);
            return l;
        }

        private void InitItem(CompoundDocumentItem item)
        {
            item.LeftSibling = -1;
            item.RightSibling = -1;
            item._handled = false;
        }

        private int AddChildren(CompoundDocumentItem item, List<CompoundDocumentItem> l)
        {
            var childId = -1;
            item.ColorFlag = 1; //Always Black-No matter here, we just add nodes as a b-tree
            if (item.Children.Count > 0)
            {
                foreach(var c in item.Children)
                {
                    InitItem(c);
                }

                item.Children.Sort();

                childId=SetSiblings(l.Count, item.Children, 0, item.Children.Count-1, -1);
                l.AddRange(item.Children);
                foreach (var c in item.Children)
                {
                    c.ChildID=AddChildren(c, l);
                }
            }
            return childId;
        }

        private void SetUnhandled(int listAdd, List<CompoundDocumentItem> children)
        {
            for(int i=0;i<children.Count;i++)
            {
                if(children[i]._handled==false)
                {
                    if(i>0 && children[i-1].RightSibling==-1 && children[i].LeftSibling!=i+listAdd-1)
                    {
                        children[i - 1].RightSibling = i + listAdd;
                    }
                    else if (i<children.Count-1 && children[i + 1].LeftSibling == -1 && children[i].RightSibling != i + listAdd+1)
                    {
                        children[i + 1].LeftSibling = i + listAdd;
                    }
                    else
                    {
                        throw (new InvalidOperationException("Invalid sibling handling in Document"));
                    }
                }
            }
        }

        private int SetSiblings(int listAdd, List<CompoundDocumentItem> children, int fromPos, int toPos, int currSibl)
        {
            int pos, div;
            pos = GetPos(fromPos,toPos);

            var item = children[pos];
            if (item._handled)
                return currSibl;
            item._handled = true;
            if (fromPos == toPos)
            {
                return fromPos + listAdd;
            }

            div = pos / 2;
            if (div <= 0)
                div = 1;            
            var lPos = GetPos(fromPos, pos-1);
            var rPos = GetPos(pos+1, toPos);
            if (div == 1 && children[lPos]._handled && children[rPos]._handled)
                return pos+ listAdd;

            if (lPos>-1 && lPos >= fromPos)
            {
                item.LeftSibling = SetSiblings(listAdd, children, fromPos, pos-1, item.LeftSibling);
            }
            if (rPos < children.Count && rPos <= toPos)
            {
                item.RightSibling = SetSiblings(listAdd, children, pos+1, toPos, item.RightSibling);
            }
            return pos + listAdd;
        }

        private int GetPos(int fromPos, int toPos)
        {
            var div=(toPos - fromPos) / 2;
            return fromPos + div;
        }

        private bool NoGreater(List<CompoundDocumentItem> children, int pos, int lPos, int listAdd)
        {
            if (pos - lPos <= 1) return true;
            for(int i=lPos+1;i<=pos; i++)
            {
                if (children[i].RightSibling!=-1 && children[i].RightSibling > lPos+ listAdd)
                    return false;
            }
            return true;
        }
        private bool NoLess(List<CompoundDocumentItem> children, int pos, int rPos, int listAdd)
        {
            if (rPos - pos <= 1) return true;
            for (int i = pos + 1; i <= rPos; i++)
            {
                if (children[i].LeftSibling != -1 && children[i].LeftSibling < rPos+ listAdd)
                    return false;
            }
            return true;
        }

        private int GetLevels(int c)
        {
            c--;
            var i = 0;
            while(c>0)
            {
                c >>=  1;
                i++;
            }
            return i;
        }

        private void FillDIFAT(BinaryWriter bw)
        {
            if (_currentDIFATSectorPos < _sectorSize)
            {
                bw.Seek(_currentDIFATSectorPos, SeekOrigin.Begin);
                while (_currentDIFATSectorPos < _sectorSize)
                {
                    if (_currentDIFATSectorPos < 512)
                    {
                        bw.Write(0xFFFFFFFF);
                    }
                    else
                    {
                        bw.Write(0x0);
                    }
                    _currentDIFATSectorPos += 4;
                }
            }
        }

        private void WritePosition(BinaryWriter bw, int sector, ref int writePos, bool isFATEntry)
        {
            int pos = (int)bw.BaseStream.Position;
            bw.Seek(writePos, SeekOrigin.Begin);
            bw.Write(sector);
            writePos += 4;
            if(isFATEntry)
            {
                CheckUpdateDIFAT(bw);
            }
            bw.Seek(pos, SeekOrigin.Begin);
        }
        private void WritePosition(BinaryWriter bw, int[] sectors, ref int writePos)
        {
            int pos = (int)bw.BaseStream.Position;
            bw.Seek(writePos, SeekOrigin.Begin);
            foreach (var sector in sectors)
            {
                bw.Write(sector);
                writePos += 4;
            }
            bw.Seek(pos, SeekOrigin.Begin);
        }
        private void WriteDirs(BinaryWriter bw, List<CompoundDocumentItem> dirs)
        {
            var miniFAT = SetMiniStream(dirs);
            AllocateFAT(bw, miniFAT.Length, dirs);
            WriteMiniFAT(bw, miniFAT);
            foreach (var entity in dirs)
            {
                if (entity.ObjectType == 5 || entity.StreamSize > _miniStreamCutoffSize)
                {
                    entity.StartingSectorLocation = WriteStream(bw, entity.Stream);
                }
            }

            WriteDirStream(bw, dirs);
        }

        private int WriteDirStream(BinaryWriter bw, List<CompoundDocumentItem> dirs)
        {
            if (dirs.Count>0)
            {
                //First dirtory sector goes into sector 2
                bw.Seek((_firstDirectorySectorLocation + 1) * _sectorSize, SeekOrigin.Begin);
                for(int i=0;i<Math.Min(_sectorSize/128,dirs.Count);i++)
                {
                    dirs[i].Write(bw);
                }
            }
            else
            {
                return -1;
            }

            bw.Seek(0, SeekOrigin.End);
            var start = (int)bw.BaseStream.Position / _sectorSize - 1;
            var pos = _sectorSize + 4;
            WritePosition(bw, start, ref pos, false);
            var streamLength = 0;
            for(int i=4;i<dirs.Count;i++)
            {
                dirs[i].Write(bw);
                streamLength += 128;
            }

            WriteStreamFullSector(bw, _sectorSize);
            WriteFAT(bw, start, streamLength);
            return start;

        }

        private void WriteMiniFAT(BinaryWriter bw, byte[] miniFAT)
        {
            if (miniFAT.Length >= _sectorSize)
            {
                bw.Seek((_firstMiniFATSectorLocation+1) * _sectorSize, SeekOrigin.Begin);
                bw.Write(miniFAT, 0, _sectorSize);
                bw.Seek(0, SeekOrigin.End);
                if (miniFAT.Length > _sectorSize)
                {
                    //Write next minifat sector to fat for sector 2
                    var sector = ((int)bw.BaseStream.Position / _sectorSize) - 1;
                    int pos = _sectorSize+(4*2);
                    WritePosition(bw, sector, ref pos, false);

                    //Write overflowing FAT sectors
                    var b = new byte[miniFAT.Length - _sectorSize];
                    Array.Copy(miniFAT, _sectorSize, b, 0, b.Length);
                    WriteStream(bw, b);
                }
                _numberofMiniFATSectors = (miniFAT.Length + 1) / _sectorSize;
            }
        }

        private int WriteStream(BinaryWriter bw, byte[] stream)
        {
            bw.Seek(0, SeekOrigin.End);
            var start = (int)bw.BaseStream.Position / _sectorSize-1;
            bw.Write(stream);
            WriteStreamFullSector(bw, _sectorSize);
            WriteFAT(bw, start, stream.Length);
            return start;
        }
        private void WriteFAT(BinaryWriter bw, int sector, long size)
        {
            bw.Seek(_currentFATSectorPos, SeekOrigin.Begin);
            var pos = _sectorSize;
            while (size > pos)
            {
                bw.Write(++sector);
                pos += _sectorSize;
                CheckUpdateDIFAT(bw);
            }
            bw.Write(END_OF_CHAIN);
            CheckUpdateDIFAT(bw);
            _currentFATSectorPos = (int)bw.BaseStream.Position;
            bw.Seek(0, SeekOrigin.End);
        }

        private void CheckUpdateDIFAT(BinaryWriter bw)
        {
            if (bw.BaseStream.Position % _sectorSize == 0)
            {
                if (_currentDIFATSectorPos % _sectorSize == 0) 
                {
                    bw.Seek(512, SeekOrigin.Current);
                }
                else if (bw.BaseStream.Position == (_sectorSize * 2))
                {
                    bw.Seek(4 * _sectorSize,SeekOrigin.Begin);    //FAT continues after initizal dir och minifat sectors.
                }
                //Add to DIFAT
                int FATSector = (int)(bw.BaseStream.Position / _sectorSize - 1);
                WritePosition(bw, FATSector, ref _currentDIFATSectorPos, false);
                _numberOfFATSectors++;
                if (_currentDIFATSectorPos == _sectorSize || ((_currentDIFATSectorPos+4)  % _sectorSize == 0 && _currentDIFATSectorPos > _sectorSize))
                {
                    bw.Write(new byte[_sectorSize]); //Write pre FAT sector
                    if (_currentDIFATSectorPos > _sectorSize)                       //Write link to next DIFAT sector
                    {
                        WritePosition(bw, FATSector+1, ref _currentDIFATSectorPos, false);
                    }
                    else
                    {
                        _firstDIFATSectorLocation = FATSector+1;                    //Current sector
                    }
                    _currentDIFATSectorPos = (int)bw.BaseStream.Position;
                    //Fill sector
                    for (int i = 0; i < _sectorSize; i++)
                    {
                        bw.Write((byte)0xFF);
                    }
                    bw.Seek(-(_sectorSize * 2), SeekOrigin.Current);
                }
            }
        }

        private void AllocateFAT(BinaryWriter bw, int miniFatLength, List<CompoundDocumentItem> dirs)
        {
            /*** First calculate full size ***/
            var fullStreamSize = (long)miniFatLength - _sectorSize; //MiniFAT starts from sector 2, by default.
            //StreamSize
            foreach (var entity in dirs)
            {
                if (entity.ObjectType == 5 || entity.StreamSize > _miniStreamCutoffSize)
                {
                    var rest = _sectorSize - entity.StreamSize % _sectorSize;
                    fullStreamSize += entity.StreamSize;
                    if (rest > 0 && rest < _sectorSize) fullStreamSize += rest;
                }
            }
            var noOfSectors = fullStreamSize / _sectorSize;

            //Directory Size
            var dirsPerSector = _sectorSize / 128;
            int dirSectors = 0;
            int firstFATSectorPos = _currentFATSectorPos;
            if (dirs.Count > dirsPerSector)
            {
                dirSectors = GetSectors(dirs.Count, dirsPerSector);
                noOfSectors += dirSectors - 1; //Four item per sector. Sector two is fixed for directories
            }

            //First calc fat no sectors and difat sectors from full size
            var numberOfFATSectors = GetSectors((int)noOfSectors, _sectorSizeInt);       //Sector 0 is already allocated
            _numberofDIFATSectors = GetDIFatSectors(numberOfFATSectors);
            noOfSectors += numberOfFATSectors + _numberofDIFATSectors;

            //Calc fat sectors again with the added fat and di fat sectors.
            numberOfFATSectors = GetSectors((int)noOfSectors, _sectorSizeInt) + _numberofDIFATSectors;
             _numberofDIFATSectors = GetDIFatSectors(numberOfFATSectors);

            //Allocate FAT and DIFAT Sectors
            bw.Write(new byte[(numberOfFATSectors + (_numberofDIFATSectors > 0 ? _numberofDIFATSectors - 1 : 0)) * _sectorSize]);

            //Move to FAT Second sector (4).
            bw.Seek(_currentFATSectorPos, SeekOrigin.Begin);
            int sectorPos = 1;
            for (int i = 1; i < 109; i++)     //We have 1 FAT sector to start with at sector 0
            {
                if (i < numberOfFATSectors + _numberofDIFATSectors)
                {
                    WriteFATItem(bw, FAT_SECTOR);
                    sectorPos++;
                }
                else
                {
                    WriteFATItem(bw, END_OF_CHAIN);
                    break;
                }
            }
            if (_numberofDIFATSectors > 0) _firstDIFATSectorLocation = sectorPos + 1;
            for (int j = 0; j < _numberofDIFATSectors; j++)
            {
                WriteFATItem(bw, DIFAT_SECTOR);
                for (int i = 0; i < _sectorSizeInt - 1; i++)
                {
                    WriteFATItem(bw, FAT_SECTOR);
                    sectorPos++;
                    if (sectorPos >= numberOfFATSectors)
                    {
                        break;
                    }
                }
                if (sectorPos > numberOfFATSectors)
                {
                    break;
                }
            }
            bw.Seek(0, SeekOrigin.End);
        }

        private int GetDIFatSectors(int FATSectors)
        {
            if (FATSectors > 109)
            {
                return GetSectors((FATSectors - 109), _sectorSizeInt-1);
            }
            else
            {
                return 0;
            }
        }

        private void WriteFATItem(BinaryWriter bw, int value)
        {
            bw.Write(value);
            CheckUpdateDIFAT(bw);
            _currentFATSectorPos = (int)bw.BaseStream.Position;            
        }

        private int GetSectors(int v, int size)
        {
            if(v % size==0)
            {
                return v / size;
            }
            else
            {
                return v / size + 1;
            }
        }

        private byte[] SetMiniStream(List<CompoundDocumentItem> dirs)
        {
            //Create the miniStream
            var ms = new MemoryStream();
            var bwMiniFATStream = new BinaryWriter(ms);
            var bwMiniFAT = new BinaryWriter(new MemoryStream());
            int pos = 0;
            foreach (var entity in dirs)
            {
                if (entity.ObjectType != 5 && entity.StreamSize>0 && entity.StreamSize <= _miniStreamCutoffSize)
                {
                    bwMiniFATStream.Write(entity.Stream);
                    WriteStreamFullSector(bwMiniFATStream, miniFATSectorSize);
                    int size = _miniSectorSize;
                    entity.StartingSectorLocation = pos;
                    while (entity.StreamSize > size)
                    {
                        bwMiniFAT.Write(++pos);
                        size += _miniSectorSize;
                    }
                    bwMiniFAT.Write(END_OF_CHAIN);
                    pos++;
                }
            }
            dirs[0].StreamSize = ms.Length;
            dirs[0].Stream = ms.ToArray();

            WriteStreamFullSector(bwMiniFAT, _sectorSize);
            return ((MemoryStream)bwMiniFAT.BaseStream).ToArray();
        }

        private static void WriteStreamFullSector(BinaryWriter bw, int sectorSize)
        {
            var rest = sectorSize - (bw.BaseStream.Length % sectorSize);
            if (rest > 0 && rest < sectorSize)
                    bw.Write(new byte[rest]);
        }
        private void WriteHeader(BinaryWriter bw)
        {
            bw.Seek(0, SeekOrigin.Begin);
            bw.Write(header);
            bw.Write(new byte[16]);             //ClsID all zero's
            bw.Write((short)0x3E);              //This field SHOULD be set to 0x003E if the major version field is either 0x0003 or 0x0004.                 
            bw.Write((short)0x3);               //Version 3
            bw.Write((ushort)0xFFFE);           // This field MUST be set to 0xFFFE. This field is a byte order mark for all integer fields, specifying little-endian byte order. 
            bw.Write((short)9);                 //Sector Shift
            bw.Write((short)6);                 //Mini Sector Shift
            bw.Write(new byte[6]);              //reserved
            bw.Write(0);                        //Number of Directory Sectors, unsupported i v3. Set to zero
            bw.Write(_numberOfFATSectors);      //Number of FAT Sectors
            bw.Write(1);                        //First Directory Sector Location
            bw.Write(0);                        //Transaction Signature Number
            bw.Write(_miniStreamCutoffSize);     //Mini Stream Cutoff Size
            bw.Write(2);                        //First Mini FAT Sector Location
            bw.Write(_numberofMiniFATSectors);   //Number of MiniFAT sectors
            bw.Write(_firstDIFATSectorLocation); //First DIFAT Sector Location
            bw.Write(_numberofDIFATSectors);     //Number of DIFAT Sectors
        }

        private void CreateFATStreams(CompoundDocumentItem item, BinaryWriter bw, BinaryWriter bwMini, DocWriteInfo dwi)
        {
            if (item.ObjectType != 5)   //Root, we must have the miniStream first.
            {
                if (item.StreamSize > 0)
                {
                    item.StreamSize = item.Stream.Length;
                    if (item.StreamSize < _miniStreamCutoffSize)
                    {
                        item.StartingSectorLocation=WriteStream(bwMini, dwi.miniFAT, item.Stream, miniFATSectorSize);
                    }
                    else
                    {
                        item.StartingSectorLocation = WriteStream(bw, dwi.FAT, item.Stream, FATSectorSizeV3);
                    }
                }
            }
            foreach(var c in item.Children)
            {
                CreateFATStreams(c, bw, bwMini, dwi);
            }
        }

        private int WriteStream(BinaryWriter bw, List<int> fat, byte[] stream, int FATSectorSize)
        {
            var rest = FATSectorSize - (stream.Length % FATSectorSize);
            bw.Write(stream);
            if(rest>0 && rest < FATSectorSize) bw.Write(new byte[rest]);
            var ret = fat.Count;
            AddFAT(fat, stream.Length, FATSectorSize, 0);

            return ret; //Returns the start sector
        }

        private void AddFAT(List<int> fat, long streamSize, int sectorSize, int addPos)
        {
            var size = 0;
            while (size<streamSize)
            {
                if (size + sectorSize < streamSize)
                {
                    fat.Add(fat.Count + 1);
                }
                else
                {
                    fat.Add(END_OF_CHAIN);
                }
                size += sectorSize;
            }
        }

        public void Dispose()
        {
            _miniSectors = null;
            _sectors = null;            
        }
#endregion
    }
}

