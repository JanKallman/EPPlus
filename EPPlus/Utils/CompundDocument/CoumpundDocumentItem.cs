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

namespace OfficeOpenXml.Utils.CompundDocument
{
    internal class CompoundDocumentItem : IComparable<CompoundDocumentItem>
    {
        public CompoundDocumentItem()
        {
            Children = new List<CompoundDocumentItem>();
        }
        public CompoundDocumentItem Parent { get; set; }
        public List<CompoundDocumentItem> Children { get; set; }

        public string Name
        {
            get;
            set;
        }

        /// <summary>
        /// 0=Red
        /// 1=Black
        /// </summary>
        public byte ColorFlag
        {
            get;
            set;
        }
        /// <summary>
        /// Type of object
        /// 0x00 - Unknown or unallocated 
        /// 0x01 - Storage Object
        /// 0x02 - Stream Object 
        /// 0x05 - Root Storage Object
        /// </summary>
        public byte ObjectType
        {
            get;

            set;
        }

        public int ChildID
        {
            get;

            set;
        }

        public Guid ClsID
        {
            get;

            set;
        }

        public int LeftSibling
        {
            get;

            set;
        }

        public int RightSibling
        {
            get;
            set;
        }

        public int StatBits
        {
            get;
            set;
        }

        public long CreationTime
        {
            get;
            set;
        }

        public long ModifiedTime
        {
            get;
            set;
        }

        public int StartingSectorLocation
        {
            get;
            set;
        }

        public long StreamSize
        {
            get;
            set;
        }

        public byte[] Stream
        {
            get;
            set;
        }
        internal bool _handled = false;
        internal void Read(BinaryReader br)
        {
            var s = br.ReadBytes(0x40);
            var sz = br.ReadInt16();
            if (sz > 0)
            {
                Name = UTF8Encoding.Unicode.GetString(s, 0, sz - 2);
            }
            ObjectType = br.ReadByte();
            ColorFlag = br.ReadByte();
            LeftSibling = br.ReadInt32();
            RightSibling = br.ReadInt32();
            ChildID = br.ReadInt32();

            //Clsid;
            ClsID = new Guid(br.ReadBytes(16));

            StatBits = br.ReadInt32();
            CreationTime = br.ReadInt64();
            ModifiedTime = br.ReadInt64();

            StartingSectorLocation = br.ReadInt32();
            StreamSize = br.ReadInt64();
        }
        internal void Write(BinaryWriter bw)
        {
            var name = Encoding.Unicode.GetBytes(Name);
            bw.Write(name);
            bw.Write(new byte[0x40 - (name.Length)]);
            bw.Write((Int16)(name.Length + 2));

            bw.Write(ObjectType);
            bw.Write(ColorFlag);
            bw.Write(LeftSibling);
            bw.Write(RightSibling);
            bw.Write(ChildID);
            bw.Write(ClsID.ToByteArray());
            bw.Write(StatBits);
            bw.Write(CreationTime);
            bw.Write(ModifiedTime);
            bw.Write(StartingSectorLocation);
            bw.Write(StreamSize);
        }

        public override string ToString()
        {
            return Name;
        }

        /// <summary>
        /// Compare length first, then sort by name in upper invariant
        /// </summary>
        /// <param name="other">The other item</param>
        /// <returns></returns>
        public int CompareTo(CompoundDocumentItem other)
        {
            if(Name.Length < other.Name.Length)
            {
                return -1;
            }
            else if(Name.Length > other.Name.Length)
            {
                return 1;
            }
            var n1 = Name.ToUpperInvariant();
            var n2 = other.Name.ToUpperInvariant();
            for (int i=0;i<n1.Length;i++)
            {
                if(n1[i] < n2[i])
                {
                    return -1;
                }
                else if(n1[i] > n2[i])
                {
                    return 1;
                }
            }
            return 0;
        }
    }
}
